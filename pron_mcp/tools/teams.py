"""Teams 쓰기 도구 (Phase 3).

Microsoft Graph API:
- 채팅 목록: GET /me/chats
- 채팅 메시지: POST /chats/{chatId}/messages
- 가입한 팀: GET /me/joinedTeams
- 채널 목록: GET /teams/{teamId}/channels
- 채널 메시지: POST /teams/{teamId}/channels/{channelId}/messages (admin 권한 필요)
- 회의 생성: POST /me/onlineMeetings

⚠️ 권한 주의:
- 1:1·그룹 채팅(`Chat.ReadWrite`)은 사용자 동의로 OK
- 팀 채널 메시지(`ChannelMessage.Send`)는 admin 동의 필요
- 회의 녹취록(`OnlineMeetingTranscript.Read.All`)도 admin 동의 필요
"""

from __future__ import annotations

from typing import Any, Literal

from mcp.server.fastmcp import Context, FastMCP
from pydantic import Field

from pron_mcp.graph_client import GraphAPIError


# ---------- 도구 등록 ----------

def register(mcp: FastMCP) -> None:
    """Teams 도구들을 mcp에 등록."""

    @mcp.tool(
        annotations={
            "title": "Teams 채팅 목록",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def list_chats(
        ctx: Context,
        top: int = Field(default=20, ge=1, le=50, description="최대 가져올 채팅 수"),
    ) -> list[dict[str, Any]]:
        """본인의 Teams 채팅 목록을 조회합니다 (1:1·그룹 채팅 포함).

        send_chat_message를 호출하려면 여기서 받은 chat_id가 필요합니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        result = await graph.get(
            "/me/chats",
            params={
                "$top": top,
                "$expand": "members",
                "$orderby": "lastUpdatedDateTime desc",
            },
        )
        return [
            {
                "chat_id": c.get("id"),
                "type": c.get("chatType"),  # oneOnOne, group, meeting
                "topic": c.get("topic"),
                "members": [
                    {
                        "email": m.get("email"),
                        "display_name": m.get("displayName"),
                    }
                    for m in c.get("members", [])
                ],
                "last_updated": c.get("lastUpdatedDateTime"),
                "web_url": c.get("webUrl"),
            }
            for c in result.get("value", [])
        ]

    @mcp.tool(
        annotations={
            "title": "Teams 채팅 메시지 전송 — 즉시 발송됨",
            "readOnlyHint": False,
            "destructiveHint": False,  # 채팅 메시지는 삭제 가능해서 destructive 아님
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def send_chat_message(
        ctx: Context,
        chat_id: str = Field(
            description=(
                "메시지를 보낼 채팅의 ID. list_chats로 받은 chat_id 사용. "
                "1:1 채팅을 새로 만들고 싶으면 먼저 create_chat 호출."
            )
        ),
        message: str = Field(description="전송할 메시지 본문"),
        body_type: Literal["text", "html"] = Field(
            default="text", description="text 또는 html"
        ),
    ) -> dict[str, Any]:
        """Teams 채팅(1:1·그룹)에 메시지를 즉시 전송합니다.

        ⚠️ 메일과 달리 초안 단계 없이 바로 발송됩니다.
        사용자 명시 승인 후에만 호출하세요. 보내는 메시지 내용을 사용자에게 보여주고 동의 확인.
        """
        graph = ctx.request_context.lifespan_context.graph
        result = await graph.post(
            f"/chats/{chat_id}/messages",
            json_body={
                "body": {"contentType": body_type, "content": message},
            },
        )
        if result is None:
            raise GraphAPIError("메시지 전송 응답이 비어있습니다")
        return {
            "status": "sent",
            "message_id": result.get("id"),
            "chat_id": chat_id,
            "sent_at": result.get("createdDateTime"),
            "web_url": result.get("webUrl"),
        }

    @mcp.tool(
        annotations={
            "title": "Teams 1:1 채팅 생성",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def create_chat(
        ctx: Context,
        recipient_email: str = Field(description="1:1 채팅 시작할 상대 이메일"),
    ) -> dict[str, Any]:
        """특정 사용자와 1:1 Teams 채팅을 시작합니다.

        이미 그 사람과 1:1 채팅이 있으면 Microsoft가 자동으로 기존 채팅을 반환합니다 (중복 안 생김).
        반환된 chat_id로 send_chat_message 호출하여 메시지 전송 가능.
        """
        graph = ctx.request_context.lifespan_context.graph
        # 본인 정보 먼저 가져와서 owner로 등록
        me = await graph.get("/me")
        my_id = me.get("id")

        result = await graph.post(
            "/chats",
            json_body={
                "chatType": "oneOnOne",
                "members": [
                    {
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "roles": ["owner"],
                        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{my_id}')",
                    },
                    {
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "roles": ["owner"],
                        "user@odata.bind": (
                            f"https://graph.microsoft.com/v1.0/users('{recipient_email}')"
                        ),
                    },
                ],
            },
        )
        if result is None:
            raise GraphAPIError("채팅 생성 응답이 비어있습니다")
        return {
            "chat_id": result.get("id"),
            "type": result.get("chatType"),
            "web_url": result.get("webUrl"),
        }

    @mcp.tool(
        annotations={
            "title": "가입한 팀 목록",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def list_joined_teams(ctx: Context) -> list[dict[str, Any]]:
        """본인이 가입한 모든 Teams 팀의 목록을 조회합니다.

        post_channel_message를 호출하려면 여기서 받은 team_id가 필요합니다.
        그 다음 list_channels로 채널 ID를 얻습니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        result = await graph.get("/me/joinedTeams")
        return [
            {
                "team_id": t.get("id"),
                "name": t.get("displayName"),
                "description": t.get("description"),
            }
            for t in result.get("value", [])
        ]

    @mcp.tool(
        annotations={
            "title": "팀 채널 목록",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def list_channels(
        ctx: Context,
        team_id: str = Field(description="채널 목록을 가져올 팀의 ID"),
    ) -> list[dict[str, Any]]:
        """특정 팀의 채널 목록을 조회합니다."""
        graph = ctx.request_context.lifespan_context.graph
        result = await graph.get(f"/teams/{team_id}/channels")
        return [
            {
                "channel_id": c.get("id"),
                "name": c.get("displayName"),
                "description": c.get("description"),
                "membership_type": c.get("membershipType"),  # standard, private, shared
                "web_url": c.get("webUrl"),
            }
            for c in result.get("value", [])
        ]

    @mcp.tool(
        annotations={
            "title": "팀 채널에 메시지 전송 — 사용자 명시 승인 후 호출 (admin 권한 필요)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def post_channel_message(
        ctx: Context,
        team_id: str = Field(description="팀 ID"),
        channel_id: str = Field(description="채널 ID"),
        message: str = Field(description="전송할 메시지 본문"),
        body_type: Literal["text", "html"] = Field(default="text"),
    ) -> dict[str, Any]:
        """팀 채널에 메시지를 게시합니다.

        ⚠️ admin 동의가 필요한 `ChannelMessage.Send` 권한 필수.
        ⚠️ 채널 메시지는 팀 전체가 볼 수 있는 공개 발언이므로, 반드시 사용자에게
        메시지 내용·대상 팀·채널을 보여주고 명시 승인 받은 후에만 호출하세요.

        권한이 부여되지 않은 경우 403 에러로 실패하며, 에러 메시지에 그 사실이 명시됩니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        result = await graph.post(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            json_body={"body": {"contentType": body_type, "content": message}},
        )
        if result is None:
            raise GraphAPIError("채널 메시지 전송 응답이 비어있습니다")
        return {
            "status": "sent",
            "message_id": result.get("id"),
            "web_url": result.get("webUrl"),
            "sent_at": result.get("createdDateTime"),
        }

    @mcp.tool(
        annotations={
            "title": "Teams 온라인 회의 생성",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def create_online_meeting(
        ctx: Context,
        subject: str = Field(description="회의 제목"),
        start: str = Field(
            description="시작 시각 (UTC 또는 timezone 포함 ISO 8601). 예: '2026-05-20T05:00:00Z' = UTC 05:00"
        ),
        end: str = Field(description="종료 시각 (UTC 또는 timezone 포함 ISO 8601)"),
        attendees: list[str] | None = Field(
            default=None, description="참석자 이메일 리스트 (회의 링크 받을 사람)"
        ),
    ) -> dict[str, Any]:
        """Teams 온라인 회의를 생성합니다 (캘린더 일정과 별개).

        반환되는 joinUrl을 메일·메시지로 공유하면 참석자가 참여 가능.
        일반적으로는 create_event에 is_online_meeting=True를 쓰는 게 더 편리합니다.
        이 도구는 캘린더 일정 없이 회의 링크만 필요한 경우용.
        """
        graph = ctx.request_context.lifespan_context.graph
        body: dict[str, Any] = {
            "startDateTime": start,
            "endDateTime": end,
            "subject": subject,
        }
        if attendees:
            body["participants"] = {
                "attendees": [
                    {
                        "identity": {
                            "user": {
                                "@odata.type": "#microsoft.graph.identity",
                            }
                        },
                        "upn": email,
                    }
                    for email in attendees
                ]
            }

        result = await graph.post("/me/onlineMeetings", json_body=body)
        if result is None:
            raise GraphAPIError("온라인 회의 생성 응답이 비어있습니다")
        return {
            "meeting_id": result.get("id"),
            "join_url": result.get("joinUrl"),
            "subject": result.get("subject"),
            "start": result.get("startDateTime"),
            "end": result.get("endDateTime"),
        }
