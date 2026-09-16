# Pron MCP

사내 업무용 **Microsoft 365 + 후이즈메일 통합 MCP 서버**.
Claude(또는 다른 MCP 클라이언트)가 회사 메일·캘린더·Teams·SharePoint·OneDrive·연락처에
접근해 실제 업무를 대신 수행할 수 있게 한다.

## 만든 의도

사내 구성원이 Claude에게 자연어로 "내일 2시 회의 잡아줘", "이 메일 초안 써줘",
"받은 메일 정리해줘" 같은 요청을 하면, Claude가 Microsoft Graph API와 후이즈메일
서버를 직접 호출해 처리하도록 만든 브리지다. 핵심 설계 목표는 세 가지다.

1. **안전한 발송 모델** — 메일·메시지처럼 되돌릴 수 없는 작업은 LLM이 단독으로
   실행하지 못하도록 "초안 생성 → 사용자 검수 → 발송"의 2단계로 분리한다.
2. **자격증명 격리** — 여러 사용자가 같은 서버에 동시 접속해도 각자의 후이즈메일
   계정이 섞이지 않도록 연결(세션)별 `ContextVar`에 자격증명을 보관한다.
3. **위임 권한 기반 인증** — client secret 없이 device code flow로 로그인한
   사용자 본인을 대신해서만 동작한다(앱 자체 권한 미사용).

## 아키텍처

```
MCP 클라이언트 (Claude Code / Claude Desktop 등)
        │  Streamable HTTP (/mcp)  ·  SSE (/sse)  ·  stdio
        ▼
  pron_mcp.server         ── ASGI 앱: /mcp, /sse, /messages/ 라우팅
        │                    요청·연결 헤더에서 후이즈 자격증명 추출 → ContextVar
        │
        ├─ [프로세스 시작 시 1회] _startup_backend()
        │     ├─ auth.OutlookAuth  ── MSAL public client + 토큰 캐시(silent 갱신)
        │     ├─ graph_client      ── Microsoft Graph API HTTP 래퍼 (Bearer 자동 주입)
        │     └─ StreamableHTTPSessionManager (stateless)
        │
        └─ tools/*           ── 도메인별 도구 등록 (메일·캘린더·Teams·파일·연락처·후이즈)
```

| 구성 요소 | 파일 | 역할 |
|---|---|---|
| 서버 엔트리 | `pron_mcp/server.py` | Streamable HTTP/SSE/stdio 전송, 세션별 자격증명 주입, 백엔드 1회 초기화 |
| 인증 | `pron_mcp/auth.py` | Entra ID 토큰 획득(캐시 silent 갱신, 0600). device code 흐름은 `allow_interactive=True`일 때만 |
| Graph 클라이언트 | `pron_mcp/graph_client.py` | Graph API 호출 + 에러를 행동 가능한 메시지로 변환 |
| 세션 저장소 | `pron_mcp/session.py` | 후이즈 자격증명용 `ContextVar` |
| 도구 모듈 | `pron_mcp/tools/` | 아래 표의 도구들 |

### 전송 방식

환경변수 `MCP_TRANSPORT`로 선택한다. 기본값 `sse`는 서버 배포용으로, 아래 **두 HTTP
엔드포인트를 동시에** 제공한다. `stdio`는 클라이언트가 프로세스를 직접 띄우는 모드다.

| 엔드포인트 | 전송 | 용도 |
|---|---|---|
| `/mcp` | **Streamable HTTP** (stateless) | 권장. `claude mcp add --transport http` 로 직접 연결(프록시 불필요) |
| `/sse` + `/messages/` | HTTP+SSE (레거시) | `mcp-remote` 등 SSE 전송을 쓰는 클라이언트용 |

`/mcp`는 stateless라 요청마다 독립 세션으로 처리되며, 자격증명 헤더가 매 요청에
실려 오므로 동시 사용자 간 격리가 자연스럽게 보장된다.

### 초기화와 인증

Graph 인증·HTTP 클라이언트·Streamable HTTP 세션 매니저는 **서버 프로세스당 1회만**
생성한다. ASGI `lifespan.startup`에서 `_startup_backend()`가 호출되어 이들을 모듈 전역에
보관하고, 연결별 `app_lifespan`은 그 공유 자원을 넘겨주기만 한다. stdio 모드에서는
첫 연결 시 지연 초기화된다.

서버는 헤드리스로 동작하므로 `get_access_token()`의 기본값은 `allow_interactive=False`다.
토큰 캐시가 만료되면 device code 흐름으로 진입하지 않고 즉시 오류를 반환한다.
재인증은 `python test_auth.py`로 수행한다.

## 보안 원칙

1. **메일 발송 2단계 분리**: `create_draft` → 사용자 검수 → `send_draft`.
   Claude는 명시 승인 없이 메일을 보내지 않는다.
2. **Teams 메시지·후이즈 발송은 즉시 전송**이므로 호출 전 사용자 확인을 거친다.
3. **파괴적 작업**(일정 삭제, 연락처 삭제 등)은 사용자 명시 승인 필수.
4. **public client 인증**: client secret을 쓰지 않는다(device code flow).
5. **토큰 캐시 권한 강화**: Unix 계열에서 사용자만 읽기/쓰기(0600).
6. **위임 권한만 사용**: 로그인한 사용자를 대신해서만 동작.

## 제공 도구

총 37개 도구를 6개 카테고리로 제공한다.

### 메일 (`tools/send_email.py`)
| 도구 | 설명 | 안전성 |
|---|---|---|
| `create_draft` | 메일 초안 생성 | 발송 안 함 |
| `list_drafts` | 초안 목록 조회 | 읽기 전용 |
| `reply_to_email` | 답장 초안 생성(스레드 유지) | 발송 안 함 |
| `forward_email` | 전달 초안 생성 | 발송 안 함 |
| `send_draft` | 초안 발송 | **명시 승인 필수** |
| `delete_draft` | 초안 삭제 | 발송 전 초안만 |

### 캘린더 (`tools/calendar.py`)
| 도구 | 설명 |
|---|---|
| `list_calendars` | 캘린더 목록 조회 |
| `create_event` | 일정 생성(Teams 회의 링크 자동 첨부 옵션) |
| `update_event` | 일정 수정·참석자 추가/제거 |
| `delete_event` | 일정 삭제(명시 승인 필수) |
| `respond_to_event` | 회의 초대 수락/거절/임시수락 |

### Teams (`tools/teams.py`)
| 도구 | 설명 |
|---|---|
| `list_chats` | 1:1·그룹 채팅 목록 |
| `create_chat` | 새 채팅 생성 |
| `send_chat_message` | 채팅 메시지 발송(확인 후) |
| `list_joined_teams` | 가입한 팀 목록 |
| `list_channels` | 팀 채널 목록 |
| `create_online_meeting` | Teams 온라인 회의 생성 |

### SharePoint · OneDrive (`tools/files.py`)
| 도구 | 설명 |
|---|---|
| `list_drives` | 드라이브 목록 |
| `list_drive_items` | 폴더 내용 조회 |
| `upload_file` | 파일 업로드(base64) |
| `create_folder` | 폴더 생성 |
| `create_sharing_link` | 공유 링크 생성 |
| `search_sharepoint_sites` | SharePoint 사이트 검색 |
| `create_sharepoint_page` | 사이트 페이지 생성 |
| `list_sharepoint_pages` | 사이트 페이지 목록 |
| `get_sharepoint_page` | 페이지 내용 조회 |
| `update_sharepoint_page` | 페이지 갱신 |
| `update_sharepoint_webpart` | 페이지 웹파트 갱신 |

### 연락처 (`tools/contacts.py`)
| 도구 | 설명 |
|---|---|
| `list_contacts` | 연락처 목록·검색 |
| `create_contact` | 연락처 생성 |
| `update_contact` | 연락처 수정 |
| `delete_contact` | 연락처 삭제(명시 승인 필수) |
| `search_people` | 조직 내 사람 검색 |

### 후이즈메일 (`tools/whois_email.py`)
회사 메일이 후이즈메일(whoisworks.com 호스팅)에 있는 경우를 위한 POP3/SMTP 연동.
계정은 요청 헤더(`X-Whois-Email`, `X-Whois-Password`)로 사용자별 전달한다.
Streamable HTTP는 매 요청마다, SSE는 연결 시점에 추출해 `ContextVar`에 보관하므로
동시 접속자끼리 계정이 섞이지 않는다. 두 헤더는 **반드시 쌍으로** 전달해야 한다
(한쪽만 오면 무시 — 짝이 깨진 자격증명으로 인증 실패하는 것을 막기 위함).

| 도구 | 설명 | 안전성 |
|---|---|---|
| `whois_list_emails` | 받은편지함 목록 조회 | 읽기 전용 |
| `whois_get_email` | 특정 메일 본문 읽기 | 읽기 전용 |
| `whois_send_email` | 메일 발송 | **명시 승인 필수** |

## 빠른 시작

자세한 단계별 안내는 [SETUP.md](SETUP.md) 참고.

```bash
# Python 3.10+
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -e .

cp .env.example .env               # AZURE_TENANT_ID / AZURE_CLIENT_ID 채우기
python test_auth.py                # 최초 1회 device code 인증 (브라우저)
python -m pron_mcp                 # 서버 실행
```

### 사전 준비 — Entra ID 앱 등록

[portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **앱 등록** → **새 등록**
(단일 테넌트, 리디렉션 URI 비움) 후:

1. **인증** → **퍼블릭 클라이언트 흐름 허용** → 예
2. **API 권한** → **Microsoft Graph** → **위임된 권한**:
   `Mail.Send`, `Mail.ReadWrite`, `Mail.Read`, `Calendars.ReadWrite`,
   `Chat.ReadWrite`, `Files.ReadWrite.All`, `Sites.ReadWrite.All`,
   `Contacts.ReadWrite`, `User.Read`, `offline_access` 등 사용할 도구에 맞게 추가
3. **개요**에서 테넌트 ID·클라이언트 ID를 `.env`에 입력

## 클라이언트 연결

서버를 한 대 띄워두고 팀원들이 각자 PC에서 붙는 구성을 가정한다.
아래 `<SERVER_HOST>`는 서버가 떠 있는 호스트(예: 사내 IP)로 바꾼다.

### 방법 A — 한 줄로 등록 (권장)

Claude Code CLI가 있으면 터미널에서 한 줄이면 끝난다. 프록시(`mcp-remote`)나
Node 설치가 필요 없다.

```bash
claude mcp add --scope user --transport http pron-mcp http://<SERVER_HOST>:8000/mcp   --header "X-Whois-Email:<본인계정>"   --header "X-Whois-Password:<본인비밀번호>"
```

- `--scope user`를 빼면 현재 폴더에서만 적용된다.
- Windows PowerShell/cmd에서는 줄바꿈(`\`) 없이 **한 줄로** 붙여넣는다.
- 비밀번호에 `^`가 있으면 cmd가 이스케이프 문자로 먹어버리므로 **방법 B**를 쓴다.

### 방법 B — 설정 파일 직접 편집 (CLI 불필요)

`~/.claude.json` (Windows: `%USERPROFILE%\.claude.json`)의 `mcpServers`에 추가:

```json
{
  "mcpServers": {
    "pron-mcp": {
      "type": "http",
      "url": "http://<SERVER_HOST>:8000/mcp",
      "headers": {
        "X-Whois-Email": "<본인계정>",
        "X-Whois-Password": "<본인비밀번호>"
      }
    }
  }
}
```

저장 후 클라이언트를 완전히 종료했다가 다시 실행한다.
비밀번호에 특수문자가 있어도 셸 이스케이프 문제가 없어 이 방법이 더 안전하다.

### 방법 C — 로컬 stdio 실행

서버를 원격에 두지 않고 클라이언트가 직접 프로세스를 띄우는 방식.
이 경우 HTTP 헤더가 없으므로 **후이즈메일 도구는 사용할 수 없다**(Graph 도구는 정상).

```json
{
  "mcpServers": {
    "pron-mcp": {
      "command": "python",
      "args": ["-m", "pron_mcp"],
      "env": {
        "MCP_TRANSPORT": "stdio",
        "AZURE_TENANT_ID": "<YOUR_TENANT_ID>",
        "AZURE_CLIENT_ID": "<YOUR_CLIENT_ID>"
      }
    }
  }
}
```

## 개발

```bash
pip install -e ".[dev]"
ruff check .
mypy pron_mcp/
pytest
```

## 라이선스

Proprietary — 사내 사용 목적. (`pyproject.toml` 참고)
</content>
