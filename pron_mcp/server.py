"""Pron MCP server entry point.

세션별 사용자 격리:
- SSE(/sse) 연결 시점 또는 Streamable HTTP(/mcp) 요청마다 헤더
  (X-Whois-Email, X-Whois-Password)에서 후이즈메일 계정 추출
- 해당 연결/요청의 ContextVar(whois_credentials)에 저장 → 도구 호출이 같은 컨텍스트에서
  실행되므로 동시 접속자끼리 자격증명이 섞이지 않는다.

두 트랜스포트를 동시에 노출한다:
- /sse + /messages/  : 구식 SSE (기존 mcp-remote 클라이언트 호환)
- /mcp               : Streamable HTTP (Claude Code 네이티브 `--transport http`)
"""

from __future__ import annotations

import logging
import os
import sys
from contextlib import AsyncExitStack, asynccontextmanager
from collections.abc import AsyncIterator
from dataclasses import dataclass

from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from mcp.server.streamable_http_manager import StreamableHTTPSessionManager

from pron_mcp.auth import OutlookAuth
from pron_mcp.graph_client import GraphClient
from pron_mcp.session import whois_credentials

logger = logging.getLogger(__name__)


@dataclass
class AppContext:
    graph: GraphClient


# 서버 프로세스당 1회만 만드는 공유 자원.
# 이전에는 SSE 연결마다 OutlookAuth/GraphClient 를 새로 만들고 매번 auth 를 돌려
# handshake 가 느려졌다(동시 접속·flapping 상황에서 클라이언트 startup timeout 유발).
_auth: OutlookAuth | None = None
_graph: GraphClient | None = None
_session_manager: StreamableHTTPSessionManager | None = None
_lifespan_stack: AsyncExitStack | None = None


async def _startup_backend() -> None:
    """ASGI startup 에서 1회 호출 — Graph 인증 + GraphClient + Streamable HTTP 매니저."""
    global _auth, _graph, _session_manager, _lifespan_stack
    if _graph is not None:
        return
    tenant_id = os.environ.get("AZURE_TENANT_ID", "").strip()
    client_id = os.environ.get("AZURE_CLIENT_ID", "").strip()
    cache_path = os.environ.get("TOKEN_CACHE_PATH")

    _auth = OutlookAuth(tenant_id=tenant_id, client_id=client_id, cache_path=cache_path)
    _graph = GraphClient(auth=_auth)
    _auth.get_access_token()  # allow_interactive=False (기본) — 캐시 무효면 즉시 실패

    # Streamable HTTP 매니저 — stateless=True로 매 요청이 독립. 자격증명 헤더가 매 요청에
    # 실려 오므로(Claude Code `--transport http --header ...`) ContextVar 격리가 자연스럽다.
    _session_manager = StreamableHTTPSessionManager(app=mcp._mcp_server, stateless=True)
    _lifespan_stack = AsyncExitStack()
    await _lifespan_stack.__aenter__()
    await _lifespan_stack.enter_async_context(_session_manager.run())

    logger.info("Pron MCP 백엔드 초기화 완료 (Graph 인증 OK, Streamable HTTP 매니저 기동, 1회)")


async def _shutdown_backend() -> None:
    global _graph, _lifespan_stack
    if _lifespan_stack is not None:
        try:
            await _lifespan_stack.__aexit__(None, None, None)
        except Exception:
            logger.exception("Streamable HTTP 매니저 종료 중 예외")
        _lifespan_stack = None
    if _graph is not None:
        await _graph.close()
        _graph = None
    logger.info("Pron MCP 백엔드 종료")


@asynccontextmanager
async def app_lifespan(_server: FastMCP) -> AsyncIterator[AppContext]:
    """연결마다 호출되지만 무거운 작업은 하지 않는다 — 공유 GraphClient만 넘긴다.

    인증/HTTP 클라이언트 생성은 _startup_backend 에서 프로세스당 1회만.
    SSE 경로는 ASGI lifespan(asgi_app)에서 먼저 호출하고, stdio 경로는 여기서
    지연 초기화로 커버한다(초기화 이후엔 no-op).
    """
    if _graph is None:
        await _startup_backend()
    yield AppContext(graph=_graph)


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


def _apply_whois_headers(scope) -> None:
    """헤더에서 X-Whois-Email/Password 를 뽑아 이 요청·연결의 ContextVar에 저장.

    Email/Password 둘 다 있을 때만 (email, password) 쌍으로 저장(짝 오염 방지).
    미설정 시 whois 도구는 명확한 에러("계정 정보 없음")를 낸다.
    """
    headers = dict(scope.get("headers", []))
    email = headers.get(b"x-whois-email", b"").decode("utf-8").strip()
    password = headers.get(b"x-whois-password", b"").decode("utf-8").strip()
    if email and password:
        whois_credentials.set((email, password))
        logger.info("후이즈메일 계정 수신: %s", email)
    elif email or password:
        logger.warning(
            "후이즈메일 자격증명 불완전 — X-Whois-Email/X-Whois-Password를 "
            "모두 전달해야 합니다. 후이즈메일 도구는 비활성됩니다."
        )


def _make_asgi_app():
    """SSE + Streamable HTTP + 세션별 자격증명 관리 ASGI 앱."""
    from mcp.server.sse import SseServerTransport

    sse_transport = SseServerTransport("/messages/")

    async def asgi_app(scope, receive, send):
        if scope["type"] == "lifespan":
            # uvicorn 프로세스 시작/종료 시 1회씩. 여기서 인증/GraphClient 를 생성한다.
            # (연결마다 반복하지 않도록 옮긴 핵심 지점)
            while True:
                message = await receive()
                if message["type"] == "lifespan.startup":
                    try:
                        await _startup_backend()
                        await send({"type": "lifespan.startup.complete"})
                    except Exception as e:
                        logger.exception("서버 startup 실패")
                        await send({"type": "lifespan.startup.failed", "message": str(e)})
                        return
                elif message["type"] == "lifespan.shutdown":
                    await _shutdown_backend()
                    await send({"type": "lifespan.shutdown.complete"})
                    return

        if scope["type"] != "http":
            return

        path = scope.get("path", "")
        method = scope.get("method", "").upper()

        if path == "/sse" and method == "GET":
            # SSE 연결 시점 헤더 → ContextVar (연결 수명 동안 유지)
            _apply_whois_headers(scope)
            async with sse_transport.connect_sse(scope, receive, send) as streams:
                await mcp._mcp_server.run(
                    streams[0], streams[1],
                    mcp._mcp_server.create_initialization_options()
                )
            return

        if path == "/messages/" or path.startswith("/messages/"):
            # POST /messages/?session_id=xxx (SSE 트랜스포트의 outgoing 경로)
            await sse_transport.handle_post_message(scope, receive, send)
            return

        if path == "/mcp" or path.startswith("/mcp"):
            # Streamable HTTP — Claude Code 네이티브 `--transport http` 경로.
            # 매 요청마다 헤더가 실려 오므로 요청별로 ContextVar 세팅.
            _apply_whois_headers(scope)
            if _session_manager is None:
                await send({"type": "http.response.start", "status": 503,
                            "headers": [[b"content-type", b"text/plain"]]})
                await send({"type": "http.response.body",
                            "body": b"session manager not initialized"})
                return
            await _session_manager.handle_request(scope, receive, send)
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
        logger.info("Pron MCP 서버 시작 — SSE: http://%s:%d/sse  |  Streamable HTTP: http://%s:%d/mcp",
                    host, port, host, port)
        app = _make_asgi_app()
        uvicorn.run(app, host=host, port=port)


if __name__ == "__main__":
    main()
