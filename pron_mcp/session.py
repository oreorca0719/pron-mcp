"""SSE 연결(세션)별 컨텍스트 저장소.

후이즈메일 자격증명은 전역 상태(os.environ)가 아니라 이 ContextVar에 담는다.
SSE 연결마다 별도의 실행 컨텍스트에서 mcp 서버가 돌고, 도구 호출도 그 컨텍스트
안에서 실행되므로(anyio start_soon이 컨텍스트를 복사) 동시 접속한 사용자끼리
자격증명이 섞이지 않는다.
"""

from __future__ import annotations

from contextvars import ContextVar

# (email, password) 쌍. 항상 둘 다 채워진 쌍으로만 set 한다 (한쪽만 들어와 짝이
# 깨지는 것을 막기 위함). 미설정 시 None.
whois_credentials: ContextVar[tuple[str, str] | None] = ContextVar(
    "whois_credentials", default=None
)
