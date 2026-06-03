"""캘린더 쓰기 도구 (Phase 2).

Microsoft Graph API 참조:
- 일정 생성: POST /me/events
- 일정 수정: PATCH /me/events/{id}
- 일정 삭제: DELETE /me/events/{id}
- 회의 응답: POST /me/events/{id}/accept|decline|tentativelyAccept
- 캘린더 목록: GET /me/calendars

기본 timezone: Asia/Seoul. 사용자가 명시하면 그것을 따른다.
"""

from __future__ import annotations

from typing import Any, Literal

from mcp.server.fastmcp import Context, FastMCP
from pydantic import Field

from pron_mcp.graph_client import GraphAPIError

DEFAULT_TIMEZONE = "Korea Standard Time"


# ---------- 공통 헬퍼 ----------

def _to_attendees(emails: list[str], type_: Literal["required", "optional"] = "required") -> list[dict[str, Any]]:
    """참석자 이메일 리스트를 Graph API attendee 객체로 변환."""
    return [
        {"emailAddress": {"address": e}, "type": type_}
        for e in emails
    ]


def _datetime_obj(dt_iso: str, tz: str) -> dict[str, str]:
    """Graph API의 dateTime 객체 형식: {dateTime, timeZone}.

    dt_iso는 'YYYY-MM-DDTHH:MM:SS' 형식의 로컬 시간 (timezone 정보 없이).
    tz는 'Korea Standard Time' 같은 Windows TZ 이름.
    """
    return {"dateTime": dt_iso, "timeZone": tz}


def _format_event_summary(event: dict[str, Any]) -> dict[str, Any]:
    """이벤트 응답을 사용자 친화적 요약으로."""
    start = event.get("start", {})
    end = event.get("end", {})
    return {
        "event_id": event.get("id"),
        "subject": event.get("subject"),
        "start": start.get("dateTime"),
        "end": end.get("dateTime"),
        "timezone": start.get("timeZone"),
        "location": (event.get("location") or {}).get("displayName"),
        "organizer": ((event.get("organizer") or {}).get("emailAddress") or {}).get("address"),
        "attendees": [
            (a.get("emailAddress") or {}).get("address")
            for a in event.get("attendees", [])
        ],
        "online_meeting_url": (event.get("onlineMeeting") or {}).get("joinUrl"),
        "web_link": event.get("webLink"),
        "is_all_day": event.get("isAllDay", False),
    }


# ---------- 도구 등록 ----------

def register(mcp: FastMCP) -> None:
    """캘린더 도구들을 mcp에 등록."""

    @mcp.tool(
        annotations={
            "title": "캘린더 목록 조회",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def list_calendars(ctx: Context) -> list[dict[str, Any]]:
        """사용자가 소유하거나 공유 받은 캘린더 목록을 조회합니다.

        create_event 호출 시 특정 캘린더에 일정 등록하려면 여기서 받은 calendar_id를 사용.
        """
        graph = ctx.request_context.lifespan_context.graph
        result = await graph.get("/me/calendars", params={"$top": 50})
        return [
            {
                "calendar_id": c.get("id"),
                "name": c.get("name"),
                "owner": (c.get("owner") or {}).get("address"),
                "is_default": c.get("isDefaultCalendar", False),
                "can_edit": c.get("canEdit", False),
                "can_share": c.get("canShare", False),
            }
            for c in result.get("value", [])
        ]

    @mcp.tool(
        annotations={
            "title": "캘린더 일정 생성",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def create_event(
        ctx: Context,
        subject: str = Field(description="일정 제목"),
        start: str = Field(
            description=(
                "시작 시각 (ISO 8601, 로컬 시간 표기). "
                "예: '2026-05-20T14:00:00' = 2026-05-20 오후 2시"
            )
        ),
        end: str = Field(
            description="종료 시각 (ISO 8601, 로컬 시간 표기). 예: '2026-05-20T15:00:00'"
        ),
        timezone: str = Field(
            default=DEFAULT_TIMEZONE,
            description=(
                "Windows timezone 이름. 한국: 'Korea Standard Time', "
                "미국 동부: 'Eastern Standard Time'"
            ),
        ),
        body: str = Field(default="", description="일정 본문/설명"),
        location: str = Field(default="", description="회의 장소"),
        attendees: list[str] | None = Field(
            default=None, description="필수 참석자 이메일 리스트"
        ),
        optional_attendees: list[str] | None = Field(
            default=None, description="선택 참석자 이메일 리스트"
        ),
        is_online_meeting: bool = Field(
            default=False,
            description="True면 Teams 회의 링크 자동 생성하여 일정에 첨부",
        ),
        is_all_day: bool = Field(default=False, description="종일 일정 여부"),
        body_type: Literal["Text", "HTML"] = Field(default="Text"),
        calendar_id: str | None = Field(
            default=None,
            description="특정 캘린더 ID. 미지정 시 기본 캘린더에 생성.",
        ),
    ) -> dict[str, Any]:
        """캘린더에 새 일정을 생성합니다.

        Teams 회의를 같이 만들고 싶으면 is_online_meeting=True로 설정. 그러면
        joinUrl이 자동 생성되어 일정 본문에 첨부되고 attendees에게 초대 메일 발송.
        """
        graph = ctx.request_context.lifespan_context.graph

        event_body: dict[str, Any] = {
            "subject": subject,
            "start": _datetime_obj(start, timezone),
            "end": _datetime_obj(end, timezone),
            "isAllDay": is_all_day,
        }
        if body:
            event_body["body"] = {"contentType": body_type, "content": body}
        if location:
            event_body["location"] = {"displayName": location}

        all_attendees = []
        if attendees:
            all_attendees.extend(_to_attendees(attendees, "required"))
        if optional_attendees:
            all_attendees.extend(_to_attendees(optional_attendees, "optional"))
        if all_attendees:
            event_body["attendees"] = all_attendees

        if is_online_meeting:
            event_body["isOnlineMeeting"] = True
            event_body["onlineMeetingProvider"] = "teamsForBusiness"

        path = f"/me/calendars/{calendar_id}/events" if calendar_id else "/me/events"
        event = await graph.post(path, json_body=event_body)
        if event is None:
            raise GraphAPIError("일정 생성 응답이 비어있습니다")
        return _format_event_summary(event)

    @mcp.tool(
        annotations={
            "title": "캘린더 일정 수정",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def update_event(
        ctx: Context,
        event_id: str = Field(description="수정할 일정의 ID"),
        subject: str | None = Field(default=None, description="새 제목 (변경 시만)"),
        start: str | None = Field(default=None, description="새 시작 시각 (변경 시만)"),
        end: str | None = Field(default=None, description="새 종료 시각 (변경 시만)"),
        timezone: str = Field(default=DEFAULT_TIMEZONE),
        body: str | None = Field(default=None, description="새 본문 (변경 시만)"),
        location: str | None = Field(default=None, description="새 장소 (변경 시만)"),
        add_attendees: list[str] | None = Field(
            default=None,
            description="추가할 참석자 (기존 + 새로 추가)",
        ),
        remove_attendees: list[str] | None = Field(
            default=None,
            description="제거할 참석자",
        ),
    ) -> dict[str, Any]:
        """기존 일정의 일부 필드를 수정합니다.

        PATCH 방식이라 명시한 필드만 변경됩니다. 참석자 추가·제거가 모두 필요하면,
        먼저 현재 참석자 목록을 가져와서 add/remove 적용 후 전체 리스트로 한 번에 갱신합니다.
        """
        graph = ctx.request_context.lifespan_context.graph

        update_body: dict[str, Any] = {}
        if subject is not None:
            update_body["subject"] = subject
        if start is not None:
            update_body["start"] = _datetime_obj(start, timezone)
        if end is not None:
            update_body["end"] = _datetime_obj(end, timezone)
        if body is not None:
            update_body["body"] = {"contentType": "Text", "content": body}
        if location is not None:
            update_body["location"] = {"displayName": location}

        # 참석자 수정은 현재 목록 가져와서 add/remove 적용
        if add_attendees or remove_attendees:
            current = await graph.get(f"/me/events/{event_id}", params={"$select": "attendees"})
            current_emails = {
                (a.get("emailAddress") or {}).get("address"): a
                for a in current.get("attendees", [])
            }
            if remove_attendees:
                for email in remove_attendees:
                    current_emails.pop(email, None)
            if add_attendees:
                for email in add_attendees:
                    if email not in current_emails:
                        current_emails[email] = {
                            "emailAddress": {"address": email},
                            "type": "required",
                        }
            update_body["attendees"] = list(current_emails.values())

        if not update_body:
            raise GraphAPIError("수정할 필드가 하나도 지정되지 않았습니다")

        result = await graph.patch(f"/me/events/{event_id}", json_body=update_body)
        if result is None:
            # 일부 PATCH는 204를 반환할 수 있음 - 재조회
            result = await graph.get(f"/me/events/{event_id}")
        return _format_event_summary(result)

    @mcp.tool(
        annotations={
            "title": "캘린더 일정 삭제 — 되돌릴 수 없음",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def delete_event(
        ctx: Context,
        event_id: str = Field(description="삭제할 일정의 ID"),
    ) -> dict[str, Any]:
        """일정을 삭제합니다.

        ⚠️ 사용자 명시 승인 없이 호출하지 마세요. 일정 삭제 시 참석자에게도 취소 알림이 발송됩니다.
        일정을 거절(decline)만 하려면 respond_to_event 사용하세요.
        """
        graph = ctx.request_context.lifespan_context.graph
        await graph.delete(f"/me/events/{event_id}")
        return {"status": "deleted", "event_id": event_id}

    @mcp.tool(
        annotations={
            "title": "회의 초대 응답",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def respond_to_event(
        ctx: Context,
        event_id: str = Field(description="응답할 회의 일정의 ID"),
        response: Literal["accept", "decline", "tentativelyAccept"] = Field(
            description="응답 종류: accept(수락), decline(거절), tentativelyAccept(임시 수락)"
        ),
        comment: str = Field(default="", description="응답에 포함할 코멘트"),
        send_response: bool = Field(
            default=True,
            description="True면 주최자에게 응답 메일 발송, False면 본인 캘린더만 갱신",
        ),
    ) -> dict[str, Any]:
        """다른 사람이 보낸 회의 초대에 응답합니다 (수락/거절/임시).

        본인 일정을 취소·삭제하려면 delete_event를 사용하세요.
        """
        graph = ctx.request_context.lifespan_context.graph
        body = {"comment": comment, "sendResponse": send_response}
        await graph.post(f"/me/events/{event_id}/{response}", json_body=body)
        return {"status": response, "event_id": event_id}
