"""인증 동작 확인용 단독 스크립트.

이 스크립트는 MCP와 무관하게 OutlookAuth + GraphClient만으로
인증이 정상 동작하는지 확인합니다.

실행:
    python test_auth.py

성공 시 출력 예시:
    ✅ 인증 성공
    👤 사용자: 홍길동 (user@example.onmicrosoft.com)
    📧 메일함 가능
"""

from __future__ import annotations

import asyncio
import os
import sys
from pathlib import Path

# 프로젝트 루트를 sys.path에 추가 (개발 환경용)
sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv

from pron_mcp.auth import OutlookAuth
from pron_mcp.graph_client import GraphClient


async def main() -> int:
    load_dotenv()

    tenant_id = os.environ.get("AZURE_TENANT_ID", "").strip()
    client_id = os.environ.get("AZURE_CLIENT_ID", "").strip()

    if not tenant_id or not client_id:
        print("❌ AZURE_TENANT_ID 또는 AZURE_CLIENT_ID가 .env에 없습니다.")
        return 1

    print(f"🔐 Tenant: {tenant_id}")
    print(f"🔐 Client: {client_id}")
    print()

    # 1) 인증 시도
    try:
        auth = OutlookAuth(tenant_id=tenant_id, client_id=client_id)
        # 단독 인증 스크립트에서만 대화형(device-code) 허용
        token = auth.get_access_token(allow_interactive=True)
        # 토큰 자체는 출력하지 않음 (보안)
        print(f"✅ 인증 성공 (토큰 길이: {len(token)} chars)")
    except Exception as e:
        print(f"❌ 인증 실패: {e}")
        return 1

    # 2) Graph API 호출 - 본인 프로필 조회
    graph = GraphClient(auth=auth)
    try:
        me = await graph.get("/me")
        display_name = me.get("displayName", "(이름 없음)")
        email = me.get("userPrincipalName", "(이메일 없음)")
        print(f"👤 사용자: {display_name} ({email})")

        # 3) 메일함 접근 확인
        folders = await graph.get("/me/mailFolders", params={"$top": 5})
        folder_count = len(folders.get("value", []))
        print(f"📧 메일함 접근 가능 (폴더 {folder_count}개+ 확인)")

        # 4) 캘린더 접근 확인
        calendars = await graph.get("/me/calendars", params={"$top": 5})
        cal_count = len(calendars.get("value", []))
        print(f"📅 캘린더 접근 가능 (캘린더 {cal_count}개+ 확인)")

    except Exception as e:
        print(f"⚠️ Graph API 호출 실패: {e}")
        return 1
    finally:
        await graph.close()

    print()
    print("🎉 모든 검증 통과! MCP 서버 실행 준비 완료.")
    print()
    print("다음 단계:")
    print("  1. python -m pron_mcp  (MCP 서버 실행 - Claude Desktop이 호출)")
    print("  2. claude_desktop_config.json에 서버 등록")
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
