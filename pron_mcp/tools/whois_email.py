"""외부 메일 계정(POP3/SMTP) 연동 도구.

계정 정보는 SSE 연결 시 헤더(X-Whois-Email, X-Whois-Password)로 전달되며
서버가 연결별 ContextVar(whois_credentials)에 저장합니다. 동시 접속자끼리
자격증명이 섞이지 않도록 전역 상태를 쓰지 않습니다.
"""

from __future__ import annotations

import os
import poplib
import smtplib
import ssl
import email as email_lib
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Any

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from pron_mcp.session import whois_credentials

DEFAULT_PORTS = {"POP3": 995, "SMTP": 587}


def _mail_server(kind: str) -> tuple[str, int]:
    """메일 서버 주소를 환경변수에서 읽는다. kind 는 "POP3" 또는 "SMTP".

    호스트는 배포 환경마다 다르므로 하드코딩하지 않고 `WHOIS_POP3_HOST`,
    `WHOIS_SMTP_HOST` 로 주입한다. 포트는 생략하면 표준 포트를 쓴다.
    """
    host = os.environ.get(f"WHOIS_{kind}_HOST", "").strip()
    if not host:
        raise ValueError(
            f"후이즈메일 {kind} 서버가 설정되지 않았습니다. 환경변수(.env)에 "
            f"WHOIS_{kind}_HOST 를 지정하세요 (예: WHOIS_{kind}_HOST=mail.example.com)."
        )
    port = int(os.environ.get(f"WHOIS_{kind}_PORT", str(DEFAULT_PORTS[kind])))
    return host, port


def _pop3_ssl_context() -> ssl.SSLContext:
    """POP3S 용 — 인증서·호스트명 검증을 모두 유지하는 기본 컨텍스트."""
    return ssl.create_default_context()


def _smtp_ssl_context() -> ssl.SSLContext:
    """SMTP STARTTLS 용 — 검증은 유지하고 암호 강도만 완화한다.

    일부 메일 서버는 작은 DH 파라미터를 사용해 OpenSSL 기본 보안 수준(SECLEVEL=2)에서
    핸드셰이크가 `DH_KEY_TOO_SMALL` 로 실패한다. 이때 필요한 최소 조치는 보안 수준을
    1로 낮추는 것뿐이며, 인증서·호스트명 검증은 그대로 유지해 중간자 공격 방어를
    포기하지 않는다.
    """
    ctx = ssl.create_default_context()
    ctx.set_ciphers("DEFAULT:@SECLEVEL=1")
    return ctx


def _decode_str(value: str | bytes | None) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        value = value.decode("utf-8", errors="replace")
    decoded_parts = decode_header(value)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            result.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(part)
    return "".join(result)


def _get_credentials() -> tuple[str, str]:
    creds = whois_credentials.get()
    if not creds or not creds[0] or not creds[1]:
        raise ValueError(
            "후이즈메일 계정 정보가 없습니다. SSE 연결 시 "
            "X-Whois-Email 과 X-Whois-Password 헤더를 함께 전달하세요 "
            "(둘 중 하나만 있으면 무시됩니다)."
        )
    return creds


def _parse_email_message(raw_lines: list[bytes]) -> dict[str, Any]:
    raw_bytes = b"\r\n".join(raw_lines)
    msg = email_lib.message_from_bytes(raw_bytes)

    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition", ""))
            if content_type == "text/plain" and "attachment" not in disposition:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    body = payload.decode(charset, errors="replace")
                    break
            elif content_type == "text/html" and not body and "attachment" not in disposition:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    body = f"[HTML 본문]\n{payload.decode(charset, errors='replace')}"
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            body = payload.decode(charset, errors="replace")

    return {
        "from": _decode_str(msg.get("From")),
        "to": _decode_str(msg.get("To")),
        "subject": _decode_str(msg.get("Subject")),
        "date": _decode_str(msg.get("Date")),
        "body": body.strip()[:3000],
    }


def _pop3_connect(email_addr: str, password: str) -> poplib.POP3_SSL:
    host, port = _mail_server("POP3")
    pop = poplib.POP3_SSL(host, port, context=_pop3_ssl_context())
    pop.user(email_addr)
    try:
        pop.pass_(password)
    except poplib.error_proto as e:
        try:
            pop.quit()
        except Exception:
            pass
        # 후이즈 게이트웨이는 계정/비밀번호가 틀려도 보안상 '-ERR internal server
        # error' 같은 일반 메시지를 돌려준다. 원인을 행동 가능하게 풀어 전달.
        raw = e.args[0] if e.args else b""
        detail = raw.decode("utf-8", "replace") if isinstance(raw, bytes) else str(raw)
        raise ValueError(
            f"후이즈메일 로그인 실패 (계정: {email_addr}). "
            f"계정 주소와 비밀번호를 확인하세요. 서버 응답: {detail}"
        ) from e
    return pop


def register(mcp: FastMCP) -> None:

    @mcp.tool(
        annotations={
            "title": "후이즈메일 받은편지함 조회",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    def whois_list_emails(
        count: int = Field(default=10, ge=1, le=50, description="가져올 최대 메일 개수 (최신순, 1-50)"),
    ) -> list[dict[str, Any]]:
        """후이즈메일 받은편지함의 최신 메일 목록을 조회합니다.
        ⚠️ 메일은 서버에서 삭제되지 않습니다 (읽기 전용 조회).
        """
        email_addr, pwd = _get_credentials()
        pop = _pop3_connect(email_addr, pwd)
        try:
            num_messages = len(pop.list()[1])
            start = num_messages
            end = max(num_messages - count, 0)
            results = []
            for i in range(start, end, -1):
                raw_lines = pop.retr(i)[1]
                parsed = _parse_email_message(raw_lines)
                results.append({
                    "index": i,
                    "from": parsed["from"],
                    "subject": parsed["subject"],
                    "date": parsed["date"],
                    "preview": parsed["body"][:150] + ("..." if len(parsed["body"]) > 150 else ""),
                })
        finally:
            try:
                pop.quit()
            except Exception:
                pass
        return results

    @mcp.tool(
        annotations={
            "title": "후이즈메일 특정 메일 본문 읽기",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    def whois_get_email(
        index: int = Field(description="읽을 메일의 번호. whois_list_emails에서 반환된 index 값."),
    ) -> dict[str, Any]:
        """후이즈메일 특정 메일의 전체 내용을 읽어옵니다.
        ⚠️ 메일은 서버에서 삭제되지 않습니다 (읽기 전용 조회).
        """
        email_addr, pwd = _get_credentials()
        pop = _pop3_connect(email_addr, pwd)
        try:
            raw_lines = pop.retr(index)[1]
        finally:
            try:
                pop.quit()
            except Exception:
                pass
        return _parse_email_message(raw_lines)

    @mcp.tool(
        annotations={
            "title": "후이즈메일로 메일 발송 — 사용자 명시 승인 후에만 호출",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    def whois_send_email(
        to: list[str] = Field(description="수신자 이메일 주소 리스트"),
        subject: str = Field(description="메일 제목"),
        body: str = Field(description="메일 본문 (텍스트)"),
        cc: list[str] | None = Field(default=None, description="참조(CC) 수신자"),
    ) -> dict[str, Any]:
        """후이즈메일 주소로 메일을 발송합니다.
        ⚠️ 절대 사용자 명시 승인 없이 호출하지 마세요.
        발송 후 회수가 불가능합니다.
        """
        email_addr, pwd = _get_credentials()

        msg = MIMEMultipart()
        msg["From"] = email_addr
        msg["To"] = ", ".join(to)
        msg["Subject"] = subject
        if cc:
            msg["Cc"] = ", ".join(cc)
        msg.attach(MIMEText(body, "plain", "utf-8"))

        bcc_self = [email_addr]
        all_recipients = to + (cc or []) + bcc_self

        host, port = _mail_server("SMTP")
        smtp = smtplib.SMTP(host, port, timeout=30)
        try:
            smtp.ehlo()
            smtp.starttls(context=_smtp_ssl_context())
            smtp.ehlo()
            smtp.login(email_addr, pwd)
            smtp.sendmail(email_addr, all_recipients, msg.as_string())
            return {
                "status": "sent",
                "from": email_addr,
                "to": to,
                "cc": cc or [],
                "subject": subject,
                "message": "후이즈메일로 메일이 발송되었습니다.",
            }
        except Exception as e:
            raise RuntimeError(f"SMTP 발송 실패: {e}") from e
        finally:
            try:
                smtp.quit()
            except Exception:
                pass
