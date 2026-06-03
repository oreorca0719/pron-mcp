"""Microsoft Entra ID authentication with token caching.

Uses MSAL Public Client + Device Code Flow:
- No client secret required (safer for desktop apps)
- User authenticates via browser with a short code
- Tokens cached to disk with strict permissions (0600)
- Automatic silent refresh when token expires
"""

from __future__ import annotations

import json
import logging
import os
import stat
import sys
from pathlib import Path

import msal

logger = logging.getLogger(__name__)

# Microsoft Graph API에 필요한 위임 권한 (사용자 대신 동작)
# 추가 권한이 필요해질 때마다 여기에 추가하고 Entra ID 앱 등록에서도 등록
GRAPH_SCOPES = [
    "Mail.Send",       # 메일 발송
    "Mail.ReadWrite",  # 메일 초안 생성·수정·삭제
    "Mail.Read",       # 메일 검색·읽기
    "User.Read",       # 사용자 프로필 조회
]

GRAPH_AUTHORITY_BASE = "https://login.microsoftonline.com"


class AuthError(Exception):
    """인증 실패."""


class OutlookAuth:
    """Microsoft Graph API 액세스 토큰 관리.

    사용 패턴:
        auth = OutlookAuth(tenant_id="...", client_id="...")
        token = auth.get_access_token()  # 자동으로 캐시·갱신·재인증 처리
    """

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        cache_path: str | Path | None = None,
    ) -> None:
        if not tenant_id or not client_id:
            raise AuthError(
                "AZURE_TENANT_ID와 AZURE_CLIENT_ID가 모두 필요합니다. "
                ".env 파일 또는 환경변수를 확인하세요."
            )

        self.tenant_id = tenant_id
        self.client_id = client_id
        self.authority = f"{GRAPH_AUTHORITY_BASE}/{tenant_id}"

        self.cache_path = self._resolve_cache_path(cache_path)
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)

        self._cache = msal.SerializableTokenCache()
        self._load_cache()

        self._app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=self._cache,
        )

    @staticmethod
    def _resolve_cache_path(cache_path: str | Path | None) -> Path:
        """캐시 경로 결정. 기본값: ~/.pron-mcp/token_cache.bin"""
        if cache_path is None:
            cache_path = "~/.pron-mcp/token_cache.bin"
        return Path(os.path.expanduser(str(cache_path)))

    def _load_cache(self) -> None:
        """디스크의 캐시 파일을 메모리로 로드."""
        if self.cache_path.exists():
            try:
                self._cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
                logger.debug("Token cache loaded from %s", self.cache_path)
            except (OSError, json.JSONDecodeError) as e:
                logger.warning("토큰 캐시 로드 실패 (재인증 필요): %s", e)

    def _save_cache(self) -> None:
        """변경사항이 있으면 디스크에 저장 + 권한 강화."""
        if not self._cache.has_state_changed:
            return
        try:
            self.cache_path.write_text(self._cache.serialize(), encoding="utf-8")
            # Unix 계열: 사용자만 읽기/쓰기 가능 (0600)
            if os.name == "posix":
                os.chmod(self.cache_path, stat.S_IRUSR | stat.S_IWUSR)
            logger.debug("Token cache saved to %s", self.cache_path)
        except OSError as e:
            logger.error("토큰 캐시 저장 실패: %s", e)

    def get_access_token(self, *, allow_interactive: bool = False) -> str:
        """유효한 액세스 토큰 반환. 캐시 → 무음 갱신 → (대화형 허용 시) device code.

        Args:
            allow_interactive: True면 캐시/무음 실패 시 device-code 흐름으로 사용자
                인증을 요청(블로킹). MCP 서버 프로세스(헤드리스 서비스)에서는 반드시
                False여야 한다 — 그렇지 않으면 device flow가 이벤트 루프를 영구
                블로킹하여 모든 연결이 멈춘다. 최초 인증은 `test_auth.py`에서만
                allow_interactive=True로 수행한다.

        Returns:
            액세스 토큰 문자열 (Bearer 헤더에 그대로 사용)

        Raises:
            AuthError: 무음 인증 실패(+대화형 비허용) 또는 모든 시도 실패
        """
        # 1) 캐시된 계정에서 무음(silent) 토큰 시도
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._save_cache()
                return result["access_token"]
            logger.info("Silent token 만료 또는 실패")

        # 헤드리스(서버) 환경: 절대 device flow로 블로킹하지 않고 즉시 명확한 에러.
        if not allow_interactive:
            raise AuthError(
                "캐시된 토큰으로 인증할 수 없습니다(만료 또는 미인증). 서버 프로세스에서는 "
                "대화형 인증을 진행하지 않습니다. 다음으로 토큰 캐시를 갱신한 뒤 서비스를 "
                "재시작하세요:\n"
                f"  (캐시 경로: {self.cache_path})\n"
                "  python test_auth.py   # 브라우저 device-code 인증 1회 수행"
            )

        # 2) Device Code Flow로 사용자 인증 요청 (allow_interactive=True 일 때만)
        flow = self._app.initiate_device_flow(scopes=GRAPH_SCOPES)
        if "user_code" not in flow:
            raise AuthError(
                f"Device Code Flow 시작 실패: {flow.get('error_description', flow)}"
            )

        # 사용자에게 인증 안내 (stderr로 출력 - MCP stdio 프로토콜 보호)
        print(
            f"\n{'=' * 60}\n"
            f"  Pron MCP 인증이 필요합니다.\n"
            f"  브라우저에서 다음 주소를 열고:\n"
            f"    {flow['verification_uri']}\n"
            f"  다음 코드를 입력하세요:\n"
            f"    {flow['user_code']}\n"
            f"{'=' * 60}\n",
            file=sys.stderr,
            flush=True,
        )

        # 사용자가 인증 완료할 때까지 대기 (블로킹)
        result = self._app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise AuthError(
                f"인증 실패: {result.get('error_description', result.get('error', 'unknown'))}"
            )

        self._save_cache()
        logger.info("인증 성공")
        return result["access_token"]

    def clear_cache(self) -> None:
        """토큰 캐시 삭제 (강제 재인증). 권한 변경 시 사용."""
        if self.cache_path.exists():
            self.cache_path.unlink()
            logger.info("토큰 캐시 삭제됨")
        # 메모리 캐시도 초기화
        self._cache = msal.SerializableTokenCache()
        self._app.token_cache = self._cache
