"""Pron MCP server entry point.

세션별 사용자 격리:
- /sse 연결 시 헤더(X-Whois-Email, X-Whois-Password)에서 후이즈메일 계정 추출
- 해당 연결의 ContextVar(whois_credentials)에 저장 → 도구 호출이 같은 컨텍스트에서
  실행되므로 동시 접속자끼리 자격증명이 섞이지 않는다.
"""

from __future__ import annotations

import logging
import os
import sys
from contextlib import asynccontextmanager
from collections.abc import AsyncIterator
from dataclasses import dataclass

from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

from pron_mcp.auth import OutlookAuth
from pron_mcp.graph_client import GraphClient
from pron_mcp.session import whois_credentials

logger = logging.getLogger(__name__)


@dataclass
class AppContext:
    graph: GraphClient


@asynccontextmanager
async def app_lifespan(_server: FastMCP) -> AsyncIterator[AppContext]:
    tenant_id = os.environ.get("AZURE_TENANT_ID", "").strip()
    client_id = os.environ.get("AZURE_CLIENT_ID", "").strip()
    cache_path = os.environ.get("TOKEN_CACHE_PATH")

    auth = OutlookAuth(
        tenant_id=tenant_id,
        client_id=client_id,
        cache_path=cache_path,
    )
    graph = GraphClient(auth=auth)

    try:
        auth.get_access_token()
        logger.info("Pron MCP 서버 인증 완료")
    except Exception as e:
        logger.error("인증 실패: %s", e)
        raise

    try:
        yield AppContext(graph=graph)
    finally:
        await graph.close()
        logger.info("Pron MCP 서버 종료")


mcp = FastMCP(
    name="pron-mcp",
    instructions=(
        "Microsoft 365 통합 MCP. 30개 도구로 메일·캘린더·Teams·SharePoint·OneDrive·연락처 작업 가능.\n\n"
        "안전 원칙:\n"
        "1) 메일: create_draft → 사용자 검수 → send_draft 2단계 분리.\n"
        "2) Teams 채팅·채널 메시지: 즉시 발송이므로 사용자 확인 후 호출.\n"
        "3) 파괴적 작업은 사용자 명시 승인 필수.\n"
        "4) ID가 필요한 도구는 먼저 list_* 또는 search_*로 ID 획득 후 호출."
    ),
    lifespan=app_lifespan,
)


def _register_tools() -> None:
    from pron_mcp.tools import calendar, contacts, files, send_email, teams, whois_email
    send_email.register(mcp)
    calendar.register(mcp)
    teams.register(mcp)
    files.register(mcp)
    contacts.register(mcp)
    whois_email.register(mcp)


def _configure_logging() -> None:
    level = os.environ.get("LOG_LEVEL", "INFO").upper()
    logging.basicConfig(
        level=getattr(logging, level, logging.INFO),
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        stream=sys.stderr,
    )


def _make_asgi_app():
    """SSE + 세션별 자격증명 관리 ASGI 앱."""
    from mcp.server.sse import SseServerTransport

    sse_transport = SseServerTransport("/messages/")

    async def asgi_app(scope, receive, send):
        if scope["type"] != "http":
            return

        path = scope.get("path", "")
        method = scope.get("method", "").upper()

        if path == "/sse" and method == "GET":
            # SSE 연결 시점: 헤더에서 자격증명 추출 → 이 연결의 ContextVar에 저장.
            # connect_sse / mcp 서버 run 이 같은 실행 컨텍스트에서 돌고, 도구 호출도
            # 그 컨텍스트에서 실행되므로(anyio가 start_soon 시 컨텍스트 복사) 동시
            # 접속자끼리 자격증명이 섞이지 않는다.
            headers = dict(scope.get("headers", []))
            email = headers.get(b"x-whois-email", b"").decode("utf-8").strip()
            password = headers.get(b"x-whois-password", b"").decode("utf-8").strip()
            if email and password:
                # 반드시 (email, password) 쌍으로 함께 저장 — 한쪽만 들어와 짝이
                # 깨지면 인증이 엉뚱하게 실패하므로(둘 다 있어야만 set).
                whois_credentials.set((email, password))
                logger.info("후이즈메일 계정 수신: %s", email)
            elif email or password:
                logger.warning(
                    "후이즈메일 자격증명 불완전 — X-Whois-Email/X-Whois-Password를 "
                    "모두 전달해야 합니다. 후이즈메일 도구는 비활성됩니다."
                )

            async with sse_transport.connect_sse(scope, receive, send) as streams:
                await mcp._mcp_server.run(
                    streams[0], streams[1],
                    mcp._mcp_server.create_initialization_options()
                )
            return

        if path == "/messages/" or path.startswith("/messages/"):
            # POST /messages/?session_id=xxx
            await sse_transport.handle_post_message(scope, receive, send)
            return

        # 404
        await send({
            "type": "http.response.start",
            "status": 404,
            "headers": [[b"content-type", b"text/plain"]],
        })
        await send({
            "type": "http.response.body",
            "body": b"Not Found",
        })

    return asgi_app


def main() -> None:
    load_dotenv()
    _configure_logging()
    _register_tools()

    transport = os.environ.get("MCP_TRANSPORT", "sse").lower()
    if transport == "stdio":
        mcp.run(transport="stdio")
    else:
        import uvicorn
        host = os.environ.get("MCP_HOST", "0.0.0.0")
        port = int(os.environ.get("MCP_PORT", "8000"))
        logger.info("Pron MCP SSE 서버 시작: http://%s:%d/sse", host, port)
        app = _make_asgi_app()
        uvicorn.run(app, host=host, port=port)


if __name__ == "__main__":
    main()
