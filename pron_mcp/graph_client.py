"""Thin wrapper around Microsoft Graph API.

Responsibilities:
- Inject Bearer token automatically (refreshes via OutlookAuth)
- Handle common error cases (401, 429, 5xx) with actionable messages
- Return parsed JSON, leaving tool-specific shaping to caller
"""

from __future__ import annotations

import logging
from typing import Any

import httpx

from pron_mcp.auth import OutlookAuth

logger = logging.getLogger(__name__)

GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_TIMEOUT = 30.0


class GraphAPIError(Exception):
    """Graph API 호출 실패. 에이전트가 다음 행동을 결정할 수 있도록
    원인을 자연어로 명확히 담는다."""

    def __init__(self, message: str, status_code: int | None = None) -> None:
        super().__init__(message)
        self.status_code = status_code


class GraphClient:
    """Microsoft Graph API HTTP 클라이언트.

    각 tool 모듈은 이 클래스의 메서드를 통해 Graph API를 호출한다.
    Bearer 토큰은 매 요청마다 OutlookAuth를 통해 자동 주입됨.
    """

    def __init__(self, auth: OutlookAuth, timeout: float = DEFAULT_TIMEOUT) -> None:
        self.auth = auth
        self.timeout = timeout
        # 비동기 클라이언트는 매 요청마다 새로 만들기보다 재사용
        self._client = httpx.AsyncClient(timeout=timeout)

    async def close(self) -> None:
        """애플리케이션 종료 시 호출."""
        await self._client.aclose()

    def _headers(self) -> dict[str, str]:
        token = self.auth.get_access_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    async def _request(
        self,
        method: str,
        path: str,
        *,
        json_body: dict[str, Any] | None = None,
        params: dict[str, Any] | None = None,
    ) -> dict[str, Any] | None:
        """공통 HTTP 호출. 경로는 '/me/messages' 같은 상대 경로."""
        url = f"{GRAPH_API_BASE}{path}"
        logger.debug("Graph API %s %s", method, path)

        try:
            response = await self._client.request(
                method=method,
                url=url,
                headers=self._headers(),
                json=json_body,
                params=params,
            )
        except httpx.RequestError as e:
            raise GraphAPIError(f"네트워크 오류: {e}") from e

        # 204 No Content (예: 메일 발송 성공)
        if response.status_code == 204:
            return None

        # 에러 응답
        if response.status_code >= 400:
            self._raise_for_status(response)

        # 빈 응답 본문 처리 (200이지만 body가 없는 경우)
        if not response.content:
            return None

        try:
            return response.json()
        except ValueError:
            # JSON 파싱 실패해도 실제 작업은 성공한 경우가 있음 (예: 메일 발송)
            # 에러를 던지지 않고 None 반환
            return None

    @staticmethod
    def _raise_for_status(response: httpx.Response) -> None:
        """Graph API 에러를 사용자가 행동 가능한 메시지로 변환."""
        status = response.status_code
        try:
            error_data = response.json().get("error", {})
            code = error_data.get("code", "Unknown")
            message = error_data.get("message", response.text[:200])
        except ValueError:
            code = "Unknown"
            message = response.text[:200]

        if status == 401:
            raise GraphAPIError(
                f"인증 실패 ({code}): {message}. "
                "토큰이 만료되었거나 권한이 부족합니다. "
                "Entra ID 앱 등록에서 위임 권한(Mail.Send, Mail.ReadWrite 등)을 "
                "확인하고 관리자 동의가 되어 있는지 확인하세요.",
                status_code=status,
            )
        if status == 403:
            raise GraphAPIError(
                f"권한 거부 ({code}): {message}. "
                "Entra ID 앱에 해당 작업 권한이 부여되지 않았습니다.",
                status_code=status,
            )
        if status == 404:
            raise GraphAPIError(
                f"리소스 없음 ({code}): {message}. ID 또는 경로를 확인하세요.",
                status_code=status,
            )
        if status == 429:
            retry_after = response.headers.get("Retry-After", "잠시 후")
            raise GraphAPIError(
                f"요청 빈도 제한 초과. {retry_after}초 후 재시도하세요.",
                status_code=status,
            )
        if 500 <= status < 600:
            raise GraphAPIError(
                f"Graph 서버 오류 ({status}): {message}. 잠시 후 재시도하세요.",
                status_code=status,
            )
        raise GraphAPIError(f"Graph API 오류 {status} ({code}): {message}", status_code=status)

    # ---- 도메인별 헬퍼 메서드 ----
    # 각 tool 파일이 직접 _request를 부르지 않고 이쪽을 통해서 호출하도록 한다.
    # 도구가 추가될수록 여기에 메서드가 늘어남.

    async def get(self, path: str, *, params: dict[str, Any] | None = None) -> dict[str, Any]:
        result = await self._request("GET", path, params=params)
        assert result is not None, "GET 응답은 비어있을 수 없음"
        return result

    async def post(
        self, path: str, *, json_body: dict[str, Any] | None = None
    ) -> dict[str, Any] | None:
        return await self._request("POST", path, json_body=json_body)

    async def patch(self, path: str, *, json_body: dict[str, Any]) -> dict[str, Any] | None:
        return await self._request("PATCH", path, json_body=json_body)

    async def delete(self, path: str) -> None:
        await self._request("DELETE", path)

    async def put_bytes(
        self,
        path: str,
        *,
        content: bytes,
        content_type: str = "application/octet-stream",
    ) -> dict[str, Any]:
        """바이너리 파일 업로드 전용. JSON이 아닌 raw bytes 본문을 보낸다.

        Graph API의 /content 엔드포인트(파일 업로드)에 사용.
        Bearer 토큰은 자동 주입되지만 Content-Type을 application/octet-stream으로
        교체해야 한다 (기본 _headers는 application/json).
        """
        url = f"{GRAPH_API_BASE}{path}"
        token = self.auth.get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": content_type,
        }
        logger.debug("Graph API PUT (bytes) %s (%d bytes)", path, len(content))

        try:
            response = await self._client.put(url, headers=headers, content=content)
        except httpx.RequestError as e:
            raise GraphAPIError(f"네트워크 오류: {e}") from e

        if response.status_code >= 400:
            self._raise_for_status(response)

        try:
            return response.json()
        except ValueError as e:
            raise GraphAPIError(f"응답 파싱 실패: {response.text[:200]}") from e
