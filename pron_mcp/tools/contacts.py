"""연락처 도구 (Phase 5).

Microsoft Graph API:
- 연락처 생성: POST /me/contacts
- 연락처 목록: GET /me/contacts
- 연락처 수정: PATCH /me/contacts/{id}
- 연락처 삭제: DELETE /me/contacts/{id}
- 사람 검색 (조직도·자주 연락하는 사람): GET /me/people

권한: Contacts.ReadWrite, People.Read (모두 사용자 동의로 가능)
"""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import Context, FastMCP
from pydantic import Field

from pron_mcp.graph_client import GraphAPIError


# ---------- 공통 헬퍼 ----------

def _format_contact(c: dict[str, Any]) -> dict[str, Any]:
    """연락처 응답을 사용자 친화적 요약으로 변환."""
    return {
        "contact_id": c.get("id"),
        "display_name": c.get("displayName"),
        "given_name": c.get("givenName"),
        "surname": c.get("surname"),
        "emails": [
            {"address": e.get("address"), "name": e.get("name")}
            for e in c.get("emailAddresses", [])
        ],
        "mobile_phone": c.get("mobilePhone"),
        "business_phones": c.get("businessPhones", []),
        "home_phones": c.get("homePhones", []),
        "job_title": c.get("jobTitle"),
        "company": c.get("companyName"),
        "department": c.get("department"),
        "office_location": c.get("officeLocation"),
        "notes": c.get("personalNotes"),
        "categories": c.get("categories", []),
        "last_modified": c.get("lastModifiedDateTime"),
    }


def _build_contact_body(
    *,
    given_name: str | None = None,
    surname: str | None = None,
    email: str | None = None,
    additional_emails: list[str] | None = None,
    mobile_phone: str | None = None,
    business_phone: str | None = None,
    job_title: str | None = None,
    company: str | None = None,
    department: str | None = None,
    office_location: str | None = None,
    notes: str | None = None,
    categories: list[str] | None = None,
) -> dict[str, Any]:
    """create/update 공통 body 구성. None인 필드는 제외."""
    body: dict[str, Any] = {}
    if given_name is not None:
        body["givenName"] = given_name
    if surname is not None:
        body["surname"] = surname

    emails = []
    if email:
        name = f"{given_name or ''} {surname or ''}".strip() or email
        emails.append({"address": email, "name": name})
    if additional_emails:
        for e in additional_emails:
            emails.append({"address": e, "name": e})
    if emails:
        body["emailAddresses"] = emails

    if mobile_phone is not None:
        body["mobilePhone"] = mobile_phone
    if business_phone is not None:
        body["businessPhones"] = [business_phone] if business_phone else []
    if job_title is not None:
        body["jobTitle"] = job_title
    if company is not None:
        body["companyName"] = company
    if department is not None:
        body["department"] = department
    if office_location is not None:
        body["officeLocation"] = office_location
    if notes is not None:
        body["personalNotes"] = notes
    if categories is not None:
        body["categories"] = categories

    return body


# ---------- 도구 등록 ----------

def register(mcp: FastMCP) -> None:
    """연락처 도구들을 mcp에 등록."""

    @mcp.tool(
        annotations={
            "title": "연락처 등록",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def create_contact(
        ctx: Context,
        given_name: str = Field(description="이름 (예: '범준')"),
        surname: str = Field(default="", description="성 (예: '김')"),
        email: str | None = Field(
            default=None, description="대표 이메일 주소 (예: 'kbj@example.com')"
        ),
        additional_emails: list[str] | None = Field(
            default=None,
            description="추가 이메일 주소 (개인용·다른 회사 주소 등)",
        ),
        mobile_phone: str | None = Field(default=None, description="휴대폰 번호"),
        business_phone: str | None = Field(default=None, description="회사 전화 번호"),
        job_title: str | None = Field(default=None, description="직책 (예: '부장')"),
        company: str | None = Field(default=None, description="회사명"),
        department: str | None = Field(default=None, description="부서"),
        office_location: str | None = Field(default=None, description="사무실 위치"),
        notes: str | None = Field(default=None, description="개인 메모"),
        categories: list[str] | None = Field(
            default=None,
            description="카테고리 태그 (예: ['고객', '협력사'])",
        ),
    ) -> dict[str, Any]:
        """Outlook 개인 연락처에 새 사람을 등록합니다.

        등록된 연락처는 Outlook 웹·앱·모바일에서 모두 동기화되어 보이고,
        메일 작성 시 수신자 자동완성으로도 사용됩니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        body = _build_contact_body(
            given_name=given_name,
            surname=surname,
            email=email,
            additional_emails=additional_emails,
            mobile_phone=mobile_phone,
            business_phone=business_phone,
            job_title=job_title,
            company=company,
            department=department,
            office_location=office_location,
            notes=notes,
            categories=categories,
        )
        result = await graph.post("/me/contacts", json_body=body)
        if result is None:
            raise GraphAPIError("연락처 생성 응답이 비어있습니다")
        return _format_contact(result)

    @mcp.tool(
        annotations={
            "title": "연락처 목록 조회",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def list_contacts(
        ctx: Context,
        search: str | None = Field(
            default=None,
            description=(
                "검색어 (이름·이메일·회사명에서 검색). "
                "예: '김부장', 'example.com', '회사명'. "
                "미지정 시 최근 수정된 순으로 전체 목록."
            ),
        ),
        top: int = Field(default=25, ge=1, le=100, description="최대 가져올 연락처 수"),
    ) -> list[dict[str, Any]]:
        """Outlook 개인 연락처 목록을 조회합니다.

        search 파라미터로 이름·이메일·회사로 필터 가능.
        update_contact나 delete_contact 호출 시 필요한 contact_id를 여기서 받습니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        params: dict[str, Any] = {
            "$top": top,
            "$orderby": "lastModifiedDateTime desc",
        }
        if search:
            # Graph API는 $search 헤더가 필요하지만 기본 $filter로 부분 검색 처리
            # displayName 기반 startswith가 가장 안정적
            params["$filter"] = (
                f"contains(displayName, '{search}') or "
                f"contains(companyName, '{search}') or "
                f"emailAddresses/any(e:contains(e/address, '{search}'))"
            )
            # 검색 시 정렬은 호환되지 않을 수 있어 제거
            params.pop("$orderby", None)

        result = await graph.get("/me/contacts", params=params)
        return [_format_contact(c) for c in result.get("value", [])]

    @mcp.tool(
        annotations={
            "title": "연락처 수정",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def update_contact(
        ctx: Context,
        contact_id: str = Field(description="수정할 연락처의 ID (list_contacts에서 받음)"),
        given_name: str | None = Field(default=None, description="새 이름 (변경 시만)"),
        surname: str | None = Field(default=None, description="새 성"),
        email: str | None = Field(default=None, description="새 대표 이메일"),
        mobile_phone: str | None = Field(default=None, description="새 휴대폰"),
        business_phone: str | None = Field(default=None, description="새 회사 전화"),
        job_title: str | None = Field(default=None, description="새 직책"),
        company: str | None = Field(default=None, description="새 회사명"),
        department: str | None = Field(default=None, description="새 부서"),
        office_location: str | None = Field(default=None, description="새 사무실"),
        notes: str | None = Field(default=None, description="새 메모"),
    ) -> dict[str, Any]:
        """기존 연락처의 일부 필드를 수정합니다.

        PATCH 방식이라 명시한 필드만 변경됩니다.
        빈 문자열을 보내면 해당 필드 삭제와 동일.
        """
        graph = ctx.request_context.lifespan_context.graph
        body = _build_contact_body(
            given_name=given_name,
            surname=surname,
            email=email,
            mobile_phone=mobile_phone,
            business_phone=business_phone,
            job_title=job_title,
            company=company,
            department=department,
            office_location=office_location,
            notes=notes,
        )
        if not body:
            raise GraphAPIError("수정할 필드가 하나도 지정되지 않았습니다")

        result = await graph.patch(f"/me/contacts/{contact_id}", json_body=body)
        if result is None:
            result = await graph.get(f"/me/contacts/{contact_id}")
        return _format_contact(result)

    @mcp.tool(
        annotations={
            "title": "연락처 삭제 — 되돌릴 수 없음",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def delete_contact(
        ctx: Context,
        contact_id: str = Field(description="삭제할 연락처의 ID"),
    ) -> dict[str, Any]:
        """연락처를 삭제합니다.

        ⚠️ 사용자 명시 승인 후에만 호출하세요. 삭제된 연락처는 복구 불가합니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        await graph.delete(f"/me/contacts/{contact_id}")
        return {"status": "deleted", "contact_id": contact_id}

    @mcp.tool(
        annotations={
            "title": "회사 사람 검색 (People API)",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def search_people(
        ctx: Context,
        query: str = Field(
            description=(
                "검색어 (이름·이메일·부서명). "
                "예: '김부장', '영업팀', '@example.com'"
            )
        ),
        top: int = Field(default=10, ge=1, le=50),
    ) -> list[dict[str, Any]]:
        """회사 디렉터리·자주 연락하는 사람·연락처를 통합 검색합니다.

        Microsoft가 사용자의 메일·회의·채팅 패턴을 학습해 관련성 높은 순서로 반환합니다.
        개인 연락처에 없어도 회사 직원이라면 결과에 포함됩니다.

        이 도구는 검색만 합니다. 회사 직원을 본인 연락처에 등록하려면
        결과의 이메일·이름을 받아 create_contact를 추가 호출하세요.
        """
        graph = ctx.request_context.lifespan_context.graph
        result = await graph.get(
            "/me/people",
            params={"$search": f'"{query}"', "$top": top},
        )
        return [
            {
                "id": p.get("id"),
                "display_name": p.get("displayName"),
                "given_name": p.get("givenName"),
                "surname": p.get("surname"),
                "job_title": p.get("jobTitle"),
                "company": p.get("companyName"),
                "department": p.get("department"),
                "office_location": p.get("officeLocation"),
                "emails": [e.get("address") for e in p.get("scoredEmailAddresses", [])],
                "phones": [ph.get("number") for ph in p.get("phones", [])],
                "person_type": (p.get("personType") or {}).get("subclass"),  # OrganizationUser, ImplicitContact 등
                "relevance_score": (p.get("scoredEmailAddresses") or [{}])[0].get("relevanceScore"),
            }
            for p in result.get("value", [])
        ]
