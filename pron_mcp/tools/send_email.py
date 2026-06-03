"""메일 쓰기 도구 (Phase 1).

안전 설계:
- create_draft / reply_to_email / forward_email은 초안만 생성. 발송 X.
- send_draft는 명시적인 발송 도구로 분리. 사용자 승인 후에만 호출.
- 도구 description에 "발송 전 사용자 검수 필수" 명시 → LLM이 자율 발송하지 않도록 유도.

Microsoft Graph API 참조:
- 초안 생성: POST /me/messages
- 초안 발송: POST /me/messages/{id}/send
- 답장 초안: POST /me/messages/{id}/createReply
- 전체답장 초안: POST /me/messages/{id}/createReplyAll
- 전달 초안: POST /me/messages/{id}/createForward
"""

from __future__ import annotations

from typing import Any, Literal

from mcp.server.fastmcp import Context, FastMCP
from pydantic import BaseModel, EmailStr, Field

from pron_mcp.graph_client import GraphAPIError


# ---------- 공통 헬퍼 ----------

def _to_recipients(emails: list[str]) -> list[dict[str, Any]]:
    """이메일 주소 리스트를 Graph API recipient 객체 리스트로 변환."""
    return [{"emailAddress": {"address": e}} for e in emails]


def _build_message_body(
    *,
    subject: str,
    body: str,
    body_type: Literal["Text", "HTML"],
    to: list[str],
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
) -> dict[str, Any]:
    """Graph API에 보낼 message JSON 본문 구성."""
    msg: dict[str, Any] = {
        "subject": subject,
        "body": {"contentType": body_type, "content": body},
        "toRecipients": _to_recipients(to),
    }
    if cc:
        msg["ccRecipients"] = _to_recipients(cc)
    if bcc:
        msg["bccRecipients"] = _to_recipients(bcc)
    return msg


def _format_draft_summary(draft: dict[str, Any]) -> dict[str, Any]:
    """발송 전 사용자 검수를 위한 초안 요약 정보."""
    return {
        "draft_id": draft.get("id"),
        "web_link": draft.get("webLink"),  # 사용자가 브라우저에서 직접 검수 가능
        "subject": draft.get("subject"),
        "to": [r["emailAddress"]["address"] for r in draft.get("toRecipients", [])],
        "cc": [r["emailAddress"]["address"] for r in draft.get("ccRecipients", [])],
        "bcc": [r["emailAddress"]["address"] for r in draft.get("bccRecipients", [])],
        "preview": (draft.get("bodyPreview") or "")[:200],
    }


# ---------- 도구 등록 ----------

def register(mcp: FastMCP) -> None:
    """pron_mcp.server에서 호출되어 도구들을 mcp 인스턴스에 등록."""

    @mcp.tool(
        annotations={
            "title": "메일 초안 생성",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def create_draft(
        ctx: Context,
        to: list[str] = Field(description="수신자 이메일 주소 리스트 (TO)"),
        subject: str = Field(description="메일 제목"),
        body: str = Field(description="메일 본문 (텍스트 또는 HTML)"),
        body_type: Literal["Text", "HTML"] = Field(
            default="Text",
            description="본문 형식. 일반 텍스트는 'Text', HTML 서식은 'HTML'",
        ),
        cc: list[str] | None = Field(default=None, description="참조(CC) 수신자"),
        bcc: list[str] | None = Field(default=None, description="숨은 참조(BCC) 수신자"),
    ) -> dict[str, Any]:
        """Outlook에 메일 초안을 생성합니다.

        ⚠️ 이 도구는 초안만 만들고 발송은 하지 않습니다.
        발송하려면 반환된 draft_id로 send_draft를 호출해야 하며,
        그 전에 반드시 사용자에게 초안 내용을 보여주고 명시적 승인을 받으세요.
        """
        graph = ctx.request_context.lifespan_context.graph
        msg = _build_message_body(
            subject=subject, body=body, body_type=body_type, to=to, cc=cc, bcc=bcc
        )
        draft = await graph.post("/me/messages", json_body=msg)
        if draft is None:
            raise GraphAPIError("초안 생성 응답이 비어있습니다")
        return _format_draft_summary(draft)

    @mcp.tool(
        annotations={
            "title": "메일 초안 목록 조회",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def list_drafts(
        ctx: Context,
        top: int = Field(
            default=10, ge=1, le=50, description="가져올 최대 초안 개수 (1-50)"
        ),
    ) -> list[dict[str, Any]]:
        """Outlook 초안 폴더에 있는 메일 초안 목록을 조회합니다.

        검수가 필요한 초안 확인, 또는 이전에 만든 초안 ID를 다시 찾을 때 사용.
        """
        graph = ctx.request_context.lifespan_context.graph
        result = await graph.get(
            "/me/mailFolders/drafts/messages",
            params={
                "$top": top,
                "$select": (
                    "id,subject,toRecipients,ccRecipients,bccRecipients,"
                    "bodyPreview,webLink,lastModifiedDateTime"
                ),
                "$orderby": "lastModifiedDateTime desc",
            },
        )
        drafts = result.get("value", [])
        return [_format_draft_summary(d) for d in drafts]

    @mcp.tool(
        annotations={
            "title": "메일 초안 발송 — 사용자 명시 승인 후에만 호출",
            "readOnlyHint": False,
            "destructiveHint": True,  # 실제 발송은 되돌릴 수 없음
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def send_draft(
        ctx: Context,
        draft_id: str = Field(
            description="발송할 초안의 ID. create_draft 또는 list_drafts에서 받은 값."
        ),
    ) -> dict[str, Any]:
        """이미 생성된 메일 초안을 즉시 발송합니다.

        ⚠️ 절대 사용자 명시 승인 없이 호출하지 마세요.
        반드시 다음 순서를 지키세요:
          1) create_draft 또는 reply_to_email로 초안 생성
          2) 초안 내용(수신자·제목·본문)을 사용자에게 보여줌
          3) 사용자가 "보내" "발송" "OK" 등 명시 승인
          4) 그 후에만 send_draft 호출

        발송 후에는 회수가 불가능합니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        # Graph API: POST /me/messages/{id}/send → 204 No Content
        await graph.post(f"/me/messages/{draft_id}/send")
        return {
            "status": "sent",
            "draft_id": draft_id,
            "message": "메일이 발송되었습니다.",
        }

    @mcp.tool(
        annotations={
            "title": "메일 답장 초안 생성",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def reply_to_email(
        ctx: Context,
        message_id: str = Field(
            description=(
                "답장할 원본 메일의 ID. M365 커넥터의 outlook_email_search로 "
                "찾은 후 read_resource의 URI에서 messageId 부분을 추출하거나, "
                "이 MCP의 향후 search 도구에서 받은 값."
            )
        ),
        body: str = Field(description="답장 본문 (원본 메일 인용은 자동으로 추가됨)"),
        reply_all: bool = Field(
            default=False,
            description="True면 전체 답장(CC 포함), False면 발신자에게만 답장",
        ),
        body_type: Literal["Text", "HTML"] = Field(default="Text"),
    ) -> dict[str, Any]:
        """기존 메일에 대한 답장 초안을 만듭니다 (스레드 유지).

        ⚠️ 초안만 생성. 발송하려면 send_draft를 별도 호출.
        Graph API의 createReply/createReplyAll은 원본 메일을 자동으로 본문에 포함하고
        제목에 'Re:' 접두사를 자동 부여합니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        endpoint = "createReplyAll" if reply_all else "createReply"
        # comment 필드로 답장 본문을 넣으면 원본 위에 자동 추가됨
        draft = await graph.post(
            f"/me/messages/{message_id}/{endpoint}",
            json_body={"comment": body},
        )
        if draft is None:
            raise GraphAPIError(f"{endpoint} 응답이 비어있습니다")

        # body_type이 HTML이면 답장 본문을 HTML로 갱신 (createReply는 기본 Text)
        if body_type == "HTML":
            await graph.patch(
                f"/me/messages/{draft['id']}",
                json_body={"body": {"contentType": "HTML", "content": body}},
            )
            # 갱신된 초안 재조회
            draft = await graph.get(f"/me/messages/{draft['id']}")

        return _format_draft_summary(draft)

    @mcp.tool(
        annotations={
            "title": "메일 전달 초안 생성",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def forward_email(
        ctx: Context,
        message_id: str = Field(description="전달할 원본 메일의 ID"),
        to: list[str] = Field(description="전달 받을 수신자 이메일"),
        comment: str = Field(
            default="",
            description="원본 위에 추가할 코멘트 (비워두면 원본만 전달)",
        ),
    ) -> dict[str, Any]:
        """기존 메일을 다른 사람에게 전달하는 초안을 만듭니다.

        ⚠️ 초안만 생성. 발송은 send_draft를 통해.
        """
        graph = ctx.request_context.lifespan_context.graph
        draft = await graph.post(
            f"/me/messages/{message_id}/createForward",
            json_body={
                "comment": comment,
                "toRecipients": _to_recipients(to),
            },
        )
        if draft is None:
            raise GraphAPIError("createForward 응답이 비어있습니다")
        return _format_draft_summary(draft)

    @mcp.tool(
        annotations={
            "title": "메일 초안 삭제",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def delete_draft(
        ctx: Context,
        draft_id: str = Field(description="삭제할 초안의 ID"),
    ) -> dict[str, Any]:
        """잘못 만든 초안을 삭제합니다 (실제 발송된 메일은 삭제 안 됨).

        발송 전 초안만 안전하게 삭제. 사용자가 원하지 않는 초안 정리용.
        """
        graph = ctx.request_context.lifespan_context.graph
        await graph.delete(f"/me/messages/{draft_id}")
        return {"status": "deleted", "draft_id": draft_id}
