# 회사 PC 셋업 가이드

## 사전 조건

- Python **3.10 이상** 설치되어 있어야 함 (`python --version` 확인)
- Windows / macOS / Linux 모두 동작
- 인터넷 접속 (Microsoft Graph API + 인증 서버)

Python이 없거나 3.9 이하라면 [python.org](https://www.python.org/downloads/)에서 최신 버전 설치.

## 1단계 — 압축 해제 & 디렉토리 진입

다운로드한 `pron-mcp-v0.1.0.zip` 압축을 풀고 터미널에서 그 디렉토리로 이동:

```bash
cd path/to/pron-mcp
```

## 2단계 — 가상환경 생성·활성화

```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
```

활성화되면 프롬프트에 `(.venv)` 가 표시됩니다.

## 3단계 — 의존성 설치

```bash
pip install -e .
```

설치되는 것: `mcp`, `msal`, `httpx`, `pydantic`, `python-dotenv`

## 4단계 — 인증 단독 테스트 (필수)

이게 가장 중요한 검증 단계입니다. MCP 통신 없이 인증만 먼저 확인:

```bash
python test_auth.py
```

처음 실행 시 stderr에 다음과 같은 화면이 나옵니다:

```
🔐 Tenant: <YOUR_TENANT_ID>
🔐 Client: <YOUR_CLIENT_ID>

============================================================
  Pron MCP 인증이 필요합니다.
  브라우저에서 다음 주소를 열고:
    https://microsoft.com/devicelogin
  다음 코드를 입력하세요:
    ABCD1234
============================================================
```

**행동 순서**:
1. 브라우저에서 `https://microsoft.com/devicelogin` 열기
2. 표시된 코드(예: `ABCD1234`) 입력 → 다음
3. 회사 계정(`user@your-tenant.onmicrosoft.com` 등)으로 로그인
4. 권한 동의 화면이 나옴 → **수락**
   - 이때 admin 동의가 안 된 권한(⚠️ 표시된 9개)은 안 보입니다 (정상)
   - 본인이 동의 가능한 약 40개 권한만 표시됨
5. 브라우저에 "디바이스에서 로그인했습니다" 메시지 표시되면 완료
6. 터미널로 돌아가면 스크립트가 진행됨

**성공 시 출력**:

```
✅ 인증 성공 (토큰 길이: 3247 chars)
👤 사용자: 홍길동 (user@your-tenant.onmicrosoft.com)
📧 메일함 접근 가능 (폴더 5개+ 확인)
📅 캘린더 접근 가능 (캘린더 1개+ 확인)

🎉 모든 검증 통과! MCP 서버 실행 준비 완료.
```

토큰은 `~/.pron-mcp/token_cache.bin` 에 저장되어 다음 실행부터는 인증 화면이 안 나옵니다.

## 5단계 — Claude Desktop에 등록

설정 파일 위치:
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

파일을 열고(없으면 생성) 다음을 추가:

```json
{
  "mcpServers": {
    "pron-mcp": {
      "command": "프로젝트경로/.venv/bin/python",
      "args": ["-m", "pron_mcp"],
      "cwd": "프로젝트경로",
      "env": {
        "AZURE_TENANT_ID": "<YOUR_TENANT_ID>",
        "AZURE_CLIENT_ID": "<YOUR_CLIENT_ID>"
      }
    }
  }
}
```

**`프로젝트경로` 자리에 실제 pron-mcp 폴더의 절대 경로**를 넣어야 합니다. 예시:

- macOS/Linux: `/Users/yourname/projects/pron-mcp`
- Windows: `C:\\Users\\yourname\\projects\\pron-mcp`
  - Windows는 `.venv/bin/python` 대신 `.venv\\Scripts\\python.exe` 사용

기존에 다른 MCP가 등록되어 있다면 `mcpServers` 객체에 `"pron-mcp"` 항목만 추가하시면 됩니다.

## 6단계 — Claude Desktop 재시작

설정 적용을 위해 Claude Desktop을 **완전히 종료** 후 재시작:
- macOS: Cmd+Q
- Windows: 시스템 트레이 우클릭 → Quit

재시작 후 새 채팅을 열어 **하단의 🔌 아이콘** 클릭하면 `pron-mcp` 서버가 보여야 합니다.

## 7단계 — 첫 사용 테스트

Claude에게 이렇게 요청해보세요:

```
내 메일 초안 목록 보여줘
```

→ `list_drafts` 호출 → 빈 목록 또는 기존 초안 출력. 여기까지 되면 성공.

다음 단계 테스트:
```
test@example.com에게 "MCP 테스트입니다" 라는 제목으로 메일 초안 만들어줘.
본문은 "안녕하세요, 이건 Pron MCP 첫 테스트입니다." 로.
```

→ `create_draft` 호출 → Outlook 웹/앱의 **초안 폴더**에서 실제 초안 확인 가능. 

발송은 명시적으로:
```
방금 만든 초안 보내줘
```

→ 사용자에게 확인을 받고 → `send_draft` 호출 → 발송 완료.

## 8단계 — Phase 2/3/4 도구 검증 시나리오

### 캘린더 (Phase 2)

```
오늘 내 캘린더 목록 보여줘
```
→ `list_calendars` 호출

```
내일 오후 2시부터 3시까지 "MCP 검증 회의" 일정 만들어줘. 
Teams 회의 링크도 같이 만들어주고.
```
→ `create_event` 호출 (`is_online_meeting=True`)

```
방금 만든 일정 3시 30분까지로 연장해줘
```
→ `update_event` 호출

```
방금 만든 일정 취소해줘
```
→ `delete_event` 호출 (사용자 확인 후)

### Teams (Phase 3)

```
내 Teams 채팅 목록 보여줘
```
→ `list_chats` 호출

```
내가 가입한 팀 목록 보여줘
```
→ `list_joined_teams` 호출 (admin 동의 없으면 일부 정보 누락 가능)

```
test@example.com한테 Teams 채팅으로 "안녕하세요" 보내줘
```
→ `create_chat` + `send_chat_message` 연쇄 (사용자 확인 후)

### SharePoint·OneDrive (Phase 4)

```
내 OneDrive 루트 폴더 내용 보여줘
```
→ `list_drives` + `list_drive_items` 호출

```
OneDrive 루트에 "MCP-테스트" 폴더 만들어줘
```
→ `create_folder` 호출

```
"테스트 내용입니다" 라는 내용으로 test.txt 파일을 OneDrive에 업로드해줘
```
→ `upload_file` 호출 (Claude가 텍스트를 base64로 인코딩 후 호출)

```
회사 SharePoint 사이트 목록 보여줘
```
→ `search_sharepoint_sites` 호출

## 문제 해결

### `ModuleNotFoundError: No module named 'mcp'`
가상환경이 활성화되지 않았거나 `pip install -e .` 가 실패함. 2~3단계 재실행.

### Device code가 안 나오고 그냥 대기
방화벽이 `login.microsoftonline.com` 접속을 막고 있을 수 있음. IT 부서 확인.

### `AADSTS50020` 또는 `AADSTS65001` 등 인증 에러
- 회사 계정이 아닌 개인 계정으로 로그인했을 가능성 → 회사 메일 계정으로 다시 시도
- 또는 권한 동의가 안 됨 → Entra 포털에서 권한 다시 확인

### Claude Desktop에 pron-mcp 서버가 안 보임
- `claude_desktop_config.json` 경로·문법 오류 (JSON 검증기로 확인)
- 절대 경로가 정확한지 확인
- Claude Desktop 로그 확인:
  - macOS: `~/Library/Logs/Claude/mcp.log`
  - Windows: `%APPDATA%\Claude\logs\mcp.log`

### `MailboxNotEnabledForRESTAPI`
해당 계정에 Exchange Online 라이선스가 없음. 메일을 보내려는 계정으로 인증해야 함.

### 권한 변경 후 도구가 작동 안 함
토큰 캐시 초기화:
```bash
rm ~/.pron-mcp/token_cache.bin   # macOS/Linux
del %USERPROFILE%\.pron-mcp\token_cache.bin   # Windows
```
다음 실행 시 재인증.
