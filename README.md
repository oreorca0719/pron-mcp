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
MCP 클라이언트 (Claude Desktop 등)
        │  SSE / stdio
        ▼
  pron_mcp.server         ── ASGI 앱: /sse, /messages/ 라우팅
        │                    SSE 연결 헤더에서 후이즈 자격증명 추출 → ContextVar
        ├─ auth.OutlookAuth  ── MSAL public client + device code flow + 토큰 캐시
        ├─ graph_client      ── Microsoft Graph API HTTP 래퍼 (Bearer 자동 주입)
        └─ tools/*           ── 도메인별 도구 등록 (메일·캘린더·Teams·파일·연락처·후이즈)
```

| 구성 요소 | 파일 | 역할 |
|---|---|---|
| 서버 엔트리 | `pron_mcp/server.py` | SSE/stdio 전송, 세션별 자격증명 주입 |
| 인증 | `pron_mcp/auth.py` | Entra ID device code flow, 토큰 캐시(0600) |
| Graph 클라이언트 | `pron_mcp/graph_client.py` | Graph API 호출 + 에러를 행동 가능한 메시지로 변환 |
| 세션 저장소 | `pron_mcp/session.py` | 후이즈 자격증명용 `ContextVar` |
| 도구 모듈 | `pron_mcp/tools/` | 아래 표의 도구들 |

전송 방식은 환경변수 `MCP_TRANSPORT`로 선택한다. 기본값 `sse`(사내 서버 배포용),
`stdio`(로컬 단독 실행용).

## 보안 원칙

1. **메일 발송 2단계 분리**: `create_draft` → 사용자 검수 → `send_draft`.
   Claude는 명시 승인 없이 메일을 보내지 않는다.
2. **Teams 메시지·후이즈 발송은 즉시 전송**이므로 호출 전 사용자 확인을 거친다.
3. **파괴적 작업**(일정 삭제, 연락처 삭제 등)은 사용자 명시 승인 필수.
4. **public client 인증**: client secret을 쓰지 않는다(device code flow).
5. **토큰 캐시 권한 강화**: Unix 계열에서 사용자만 읽기/쓰기(0600).
6. **위임 권한만 사용**: 로그인한 사용자를 대신해서만 동작.

## 제공 도구

총 36개 도구를 6개 카테고리로 제공한다.

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
계정은 SSE 연결 헤더(`X-Whois-Email`, `X-Whois-Password`)로 사용자별 전달한다.

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

### Claude Desktop 연결

`claude_desktop_config.json` (macOS: `~/Library/Application Support/Claude/`,
Windows: `%APPDATA%\Claude\`):

```json
{
  "mcpServers": {
    "pron-mcp": {
      "command": "python",
      "args": ["-m", "pron_mcp"],
      "env": {
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
