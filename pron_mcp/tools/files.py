"""SharePoint·OneDrive 파일 쓰기 도구 (Phase 4).

Microsoft Graph API:
- OneDrive 업로드: PUT /me/drive/root:/{path}:/content
- SharePoint 업로드: PUT /sites/{siteId}/drive/root:/{path}:/content
- 폴더 생성: POST /me/drive/root/children
- 공유 링크: POST /me/drive/items/{id}/createLink
- SharePoint 페이지: POST /sites/{siteId}/pages
- SharePoint 페이지 조회·수정: GET/PATCH /sites/{siteId}/pages/{pageId}/microsoft.graph.sitePage
- SharePoint webPart 수정: PATCH /sites/{siteId}/pages/{pageId}/microsoft.graph.sitePage/webParts/{webpartId}

⚠️ 파일 크기 제한:
- 이 도구는 4MB 이하 simple upload만 지원
- 4MB 초과는 chunked upload 세션 필요 (향후 추가)
"""

from __future__ import annotations

import base64
from typing import Any, Literal

from mcp.server.fastmcp import Context, FastMCP
from pydantic import Field

from pron_mcp.graph_client import GraphAPIError

SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024  # 4 MB


# ---------- 도구 등록 ----------

def register(mcp: FastMCP) -> None:
    """파일 도구들을 mcp에 등록."""

    @mcp.tool(
        annotations={
            "title": "본인의 OneDrive·SharePoint 드라이브 목록",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def list_drives(ctx: Context) -> list[dict[str, Any]]:
        """접근 가능한 OneDrive·SharePoint 드라이브 목록을 조회합니다.

        본인 OneDrive와 가입한 그룹·사이트의 문서 라이브러리 모두 포함.
        """
        graph = ctx.request_context.lifespan_context.graph
        # 본인 OneDrive
        my_drive = await graph.get("/me/drive")
        result = [
            {
                "drive_id": my_drive.get("id"),
                "name": my_drive.get("name") or "OneDrive",
                "type": "personal",
                "owner": ((my_drive.get("owner") or {}).get("user") or {}).get("displayName"),
                "web_url": my_drive.get("webUrl"),
                "is_my_drive": True,
            }
        ]
        # 본인이 접근 가능한 그룹·사이트 드라이브
        shared = await graph.get("/me/drives")
        for d in shared.get("value", []):
            if d.get("id") == my_drive.get("id"):
                continue
            result.append(
                {
                    "drive_id": d.get("id"),
                    "name": d.get("name"),
                    "type": d.get("driveType"),  # business, personal, documentLibrary
                    "owner": ((d.get("owner") or {}).get("user") or {}).get("displayName"),
                    "web_url": d.get("webUrl"),
                    "is_my_drive": False,
                }
            )
        return result

    @mcp.tool(
        annotations={
            "title": "드라이브 안의 폴더·파일 목록",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def list_drive_items(
        ctx: Context,
        path: str = Field(
            default="/",
            description=(
                "조회할 폴더 경로. '/' = 드라이브 루트. 예: '/문서', '/Projects/2026'"
            ),
        ),
        drive_id: str | None = Field(
            default=None,
            description="특정 드라이브 ID. 미지정 시 본인 OneDrive.",
        ),
        top: int = Field(default=50, ge=1, le=200),
    ) -> list[dict[str, Any]]:
        """드라이브의 특정 폴더 내용을 조회합니다."""
        graph = ctx.request_context.lifespan_context.graph
        base = f"/drives/{drive_id}" if drive_id else "/me/drive"

        # 루트는 경로 표기가 다름
        if path == "/" or path == "":
            endpoint = f"{base}/root/children"
        else:
            # 경로 인코딩
            clean_path = path.strip("/")
            endpoint = f"{base}/root:/{clean_path}:/children"

        result = await graph.get(endpoint, params={"$top": top})
        return [
            {
                "item_id": item.get("id"),
                "name": item.get("name"),
                "type": "folder" if "folder" in item else "file",
                "size_bytes": item.get("size"),
                "last_modified": item.get("lastModifiedDateTime"),
                "web_url": item.get("webUrl"),
                "download_url": item.get("@microsoft.graph.downloadUrl") if "file" in item else None,
            }
            for item in result.get("value", [])
        ]

    @mcp.tool(
        annotations={
            "title": "OneDrive·SharePoint 파일 업로드 (4MB 이하)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def upload_file(
        ctx: Context,
        path: str = Field(
            description=(
                "업로드할 파일의 드라이브 경로. "
                "예: '/문서/회의록.txt', '/Projects/2026/proposal.docx'"
            )
        ),
        content_base64: str = Field(
            description=(
                "Base64로 인코딩된 파일 내용. "
                "텍스트 파일은 UTF-8로 인코딩 후 Base64로 변환. "
                "이미 base64 형태의 첨부 데이터가 있으면 그대로 전달."
            )
        ),
        content_type: str = Field(
            default="application/octet-stream",
            description=(
                "파일의 MIME 타입. "
                "예: 'text/plain', 'application/pdf', "
                "'application/vnd.openxmlformats-officedocument.wordprocessingml.document' (docx)"
            ),
        ),
        drive_id: str | None = Field(
            default=None,
            description="업로드할 드라이브 ID. 미지정 시 본인 OneDrive.",
        ),
        conflict_behavior: Literal["rename", "replace", "fail"] = Field(
            default="rename",
            description=(
                "같은 이름 파일이 있을 때: rename(자동 이름변경, 안전), "
                "replace(기존 덮어쓰기, ⚠️ 데이터 손실 가능), fail(에러)"
            ),
        ),
    ) -> dict[str, Any]:
        """OneDrive 또는 SharePoint에 파일을 업로드합니다.

        ⚠️ 4MB 이하 파일만 지원 (큰 파일은 향후 chunked upload 추가 예정).
        ⚠️ conflict_behavior='replace'는 기존 파일을 덮어씁니다 - 사용자 명시 승인 필수.

        텍스트 파일 업로드 예시 (Python):
            text = "회의록 내용..."
            content_base64 = base64.b64encode(text.encode('utf-8')).decode()
        """
        graph = ctx.request_context.lifespan_context.graph

        # Base64 디코드
        try:
            content_bytes = base64.b64decode(content_base64)
        except Exception as e:
            raise GraphAPIError(f"Base64 디코딩 실패: {e}") from e

        if len(content_bytes) > SIMPLE_UPLOAD_LIMIT:
            raise GraphAPIError(
                f"파일 크기({len(content_bytes):,} bytes)가 4MB 제한을 초과합니다. "
                "큰 파일은 chunked upload가 필요한데 이 버전에서는 미지원입니다."
            )

        clean_path = path.strip("/")
        base = f"/drives/{drive_id}" if drive_id else "/me/drive"
        # @microsoft.graph.conflictBehavior는 쿼리 파라미터로
        endpoint = (
            f"{base}/root:/{clean_path}:/content"
            f"?@microsoft.graph.conflictBehavior={conflict_behavior}"
        )

        result = await graph.put_bytes(
            endpoint, content=content_bytes, content_type=content_type
        )
        return {
            "status": "uploaded",
            "item_id": result.get("id"),
            "name": result.get("name"),
            "size_bytes": result.get("size"),
            "web_url": result.get("webUrl"),
            "path": (result.get("parentReference") or {}).get("path"),
        }

    @mcp.tool(
        annotations={
            "title": "폴더 생성",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def create_folder(
        ctx: Context,
        parent_path: str = Field(
            default="/",
            description="부모 폴더 경로. '/' = 드라이브 루트.",
        ),
        folder_name: str = Field(description="새 폴더 이름"),
        drive_id: str | None = Field(default=None, description="드라이브 ID. 미지정 시 본인 OneDrive."),
    ) -> dict[str, Any]:
        """드라이브에 새 폴더를 생성합니다."""
        graph = ctx.request_context.lifespan_context.graph
        base = f"/drives/{drive_id}" if drive_id else "/me/drive"

        if parent_path == "/" or parent_path == "":
            endpoint = f"{base}/root/children"
        else:
            clean_path = parent_path.strip("/")
            endpoint = f"{base}/root:/{clean_path}:/children"

        result = await graph.post(
            endpoint,
            json_body={
                "name": folder_name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "rename",
            },
        )
        if result is None:
            raise GraphAPIError("폴더 생성 응답이 비어있습니다")
        return {
            "folder_id": result.get("id"),
            "name": result.get("name"),
            "web_url": result.get("webUrl"),
        }

    @mcp.tool(
        annotations={
            "title": "파일·폴더 공유 링크 생성",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def create_sharing_link(
        ctx: Context,
        item_id: str = Field(description="공유할 파일·폴더의 ID"),
        link_type: Literal["view", "edit"] = Field(
            default="view",
            description="view(읽기 전용) 또는 edit(편집 가능)",
        ),
        scope: Literal["organization", "anonymous"] = Field(
            default="organization",
            description=(
                "organization(회사 직원만 접근), anonymous(링크 아는 모두). "
                "⚠️ 금융권 보안 정책상 anonymous는 회사 정책으로 막혀있을 수 있음."
            ),
        ),
        drive_id: str | None = Field(default=None, description="파일이 있는 드라이브 ID"),
    ) -> dict[str, Any]:
        """파일·폴더의 공유 링크를 생성합니다.

        ⚠️ 외부 공유(anonymous)는 사용자 명시 승인 필수. 회사 보안 정책에 따라
        anonymous 옵션이 거부될 수 있으며, 그 경우 403 에러로 실패합니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        base = f"/drives/{drive_id}" if drive_id else "/me/drive"
        result = await graph.post(
            f"{base}/items/{item_id}/createLink",
            json_body={"type": link_type, "scope": scope},
        )
        if result is None:
            raise GraphAPIError("공유 링크 생성 응답이 비어있습니다")
        return {
            "share_url": (result.get("link") or {}).get("webUrl"),
            "link_type": link_type,
            "scope": scope,
            "expiration": result.get("expirationDateTime"),
        }

    @mcp.tool(
        annotations={
            "title": "SharePoint 사이트 검색",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def search_sharepoint_sites(
        ctx: Context,
        query: str = Field(
            default="*",
            description="사이트 이름 또는 키워드. '*' = 모든 사이트",
        ),
    ) -> list[dict[str, Any]]:
        """접근 가능한 SharePoint 사이트를 검색합니다.

        create_sharepoint_page를 호출하려면 여기서 받은 site_id가 필요합니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        result = await graph.get("/sites", params={"search": query})
        return [
            {
                "site_id": s.get("id"),
                "name": s.get("displayName"),
                "description": s.get("description"),
                "web_url": s.get("webUrl"),
                "created": s.get("createdDateTime"),
            }
            for s in result.get("value", [])
        ]

    @mcp.tool(
        annotations={
            "title": "SharePoint 페이지 생성 (사내 위키)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def create_sharepoint_page(
        ctx: Context,
        site_id: str = Field(description="페이지를 만들 SharePoint 사이트 ID"),
        title: str = Field(description="페이지 제목"),
        html_content: str = Field(
            description=(
                "페이지 본문 (HTML). 예: '<h1>제목</h1><p>본문</p>'. "
                "마크다운은 미지원이므로 사전 변환 필요."
            )
        ),
        publish: bool = Field(
            default=False,
            description=(
                "True면 즉시 게시(전사 노출), False면 초안으로 저장. "
                "⚠️ publish=True는 사용자 명시 승인 필수"
            ),
        ),
    ) -> dict[str, Any]:
        """SharePoint 사이트에 새 페이지를 생성합니다 (사내 위키, 공지 등).

        publish=False면 초안만 만들고 사용자가 검토 후 SharePoint UI에서 직접 게시 가능.
        publish=True는 즉시 사이트 구성원 전원에게 노출되므로 신중히 호출.
        """
        graph = ctx.request_context.lifespan_context.graph

        # 페이지 생성 (modern page format)
        page_body = {
            "@odata.type": "#microsoft.graph.sitePage",
            "name": f"{title.replace(' ', '-')}.aspx",
            "title": title,
            "pageLayout": "article",
            "showComments": True,
            "showRecommendedPages": False,
            "titleArea": {
                "enableGradientEffect": False,
                "imageWebUrl": None,
                "layout": "plain",
                "showAuthor": True,
                "showPublishedDate": True,
                "showTextBlockAboveTitle": False,
                "textAboveTitle": "",
                "textAlignment": "left",
                "title": title,
            },
            "canvasLayout": {
                "horizontalSections": [
                    {
                        "layout": "oneColumn",
                        "id": "1",
                        "emphasis": "none",
                        "columns": [
                            {
                                "id": "1",
                                "width": 12,
                                "webparts": [
                                    {
                                        "id": "text-1",
                                        "innerHtml": html_content,
                                    }
                                ],
                            }
                        ],
                    }
                ]
            },
        }

        page = await graph.post(f"/sites/{site_id}/pages", json_body=page_body)
        if page is None:
            raise GraphAPIError("페이지 생성 응답이 비어있습니다")

        page_id = page.get("id")
        result: dict[str, Any] = {
            "page_id": page_id,
            "name": page.get("name"),
            "title": page.get("title"),
            "web_url": page.get("webUrl"),
            "status": "draft",
        }

        # 게시
        if publish:
            await graph.post(f"/sites/{site_id}/pages/{page_id}/microsoft.graph.sitePage/publish")
            result["status"] = "published"

        return result

    @mcp.tool(
        annotations={
            "title": "SharePoint 사이트의 페이지 목록 조회",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def list_sharepoint_pages(
        ctx: Context,
        site_id: str = Field(description="페이지 목록을 가져올 SharePoint 사이트 ID"),
    ) -> list[dict[str, Any]]:
        """SharePoint 사이트의 페이지 목록을 조회합니다.

        get_sharepoint_page, update_sharepoint_page, update_sharepoint_webpart 등을
        호출하려면 여기서 받은 page_id가 필요합니다.
        """
        graph = ctx.request_context.lifespan_context.graph
        result = await graph.get(f"/sites/{site_id}/pages")
        return [
            {
                "page_id": p.get("id"),
                "name": p.get("name"),
                "title": p.get("title"),
                "web_url": p.get("webUrl"),
                "published": (p.get("publishingState") or {}).get("level") == "published",
                "last_modified": p.get("lastModifiedDateTime"),
            }
            for p in result.get("value", [])
        ]

    @mcp.tool(
        annotations={
            "title": "SharePoint 페이지 본문·구조 조회",
            "readOnlyHint": True,
            "idempotentHint": True,
            "openWorldHint": True,
        }
    )
    async def get_sharepoint_page(
        ctx: Context,
        site_id: str = Field(description="SharePoint 사이트 ID"),
        page_id: str = Field(
            description="조회할 페이지 ID (list_sharepoint_pages에서 받음)"
        ),
    ) -> dict[str, Any]:
        """SharePoint 페이지의 메타데이터와 본문 webPart들을 조회합니다.

        응답의 `webparts` 리스트에서 각 webPart의 `webpart_id`와 `inner_html`(텍스트
        타입) 또는 `web_part_type`(표준 타입)을 확인한 뒤, update_sharepoint_webpart를
        호출해 특정 webPart의 내용만 교체할 수 있습니다.

        정보 최신화 워크플로우:
        1. get_sharepoint_page로 현재 본문 구조와 내용 조회
        2. 어떤 webPart의 내용이 stale인지 식별
        3. 사용자와 새 내용 확정
        4. update_sharepoint_webpart로 해당 webPart의 innerHtml 교체
        """
        graph = ctx.request_context.lifespan_context.graph
        page = await graph.get(
            f"/sites/{site_id}/pages/{page_id}/microsoft.graph.sitePage",
            params={"$expand": "canvasLayout"},
        )

        # webPart를 평탄 리스트로 추출 (한눈에 보고 수정 대상 고르기 쉽게)
        webparts: list[dict[str, Any]] = []
        canvas = page.get("canvasLayout") or {}
        for section in canvas.get("horizontalSections") or []:
            section_id = section.get("id")
            for column in section.get("columns") or []:
                column_id = column.get("id")
                for wp in column.get("webparts") or []:
                    entry: dict[str, Any] = {
                        "webpart_id": wp.get("id"),
                        "section_id": section_id,
                        "column_id": column_id,
                    }
                    # textWebPart는 innerHtml, standardWebPart는 webPartType + data
                    if "innerHtml" in wp:
                        entry["type"] = "text"
                        entry["inner_html"] = wp.get("innerHtml")
                    else:
                        entry["type"] = "standard"
                        entry["web_part_type"] = wp.get("webPartType")
                        entry["data"] = wp.get("data")
                    webparts.append(entry)

        return {
            "page_id": page.get("id"),
            "name": page.get("name"),
            "title": page.get("title"),
            "web_url": page.get("webUrl"),
            "published": (page.get("publishingState") or {}).get("level") == "published",
            "last_modified": page.get("lastModifiedDateTime"),
            "webparts": webparts,
        }

    @mcp.tool(
        annotations={
            "title": "SharePoint 페이지 메타데이터 수정 (제목 등)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def update_sharepoint_page(
        ctx: Context,
        site_id: str = Field(description="SharePoint 사이트 ID"),
        page_id: str = Field(description="수정할 페이지 ID"),
        title: str | None = Field(
            default=None, description="새 페이지 제목 (변경 시만)"
        ),
    ) -> dict[str, Any]:
        """SharePoint 페이지의 메타데이터(제목 등)를 수정합니다.

        PATCH 방식이라 명시한 필드만 변경됩니다. 본문 내용을 바꾸려면
        update_sharepoint_webpart를 사용하세요 (이 도구로는 본문 수정 불가).
        """
        if title is None:
            raise GraphAPIError(
                "최소 한 개 이상의 변경 필드를 지정하세요 (현재 지원: title)"
            )

        graph = ctx.request_context.lifespan_context.graph
        body: dict[str, Any] = {"@odata.type": "#microsoft.graph.sitePage"}
        if title is not None:
            body["title"] = title

        updated = await graph.patch(
            f"/sites/{site_id}/pages/{page_id}/microsoft.graph.sitePage",
            json_body=body,
        )
        if updated is None:
            raise GraphAPIError("페이지 메타데이터 수정 응답이 비어있습니다")

        return {
            "page_id": updated.get("id"),
            "name": updated.get("name"),
            "title": updated.get("title"),
            "web_url": updated.get("webUrl"),
            "last_modified": updated.get("lastModifiedDateTime"),
        }

    @mcp.tool(
        annotations={
            "title": "SharePoint 페이지 webPart 본문 수정 (정보 최신화)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        }
    )
    async def update_sharepoint_webpart(
        ctx: Context,
        site_id: str = Field(description="SharePoint 사이트 ID"),
        page_id: str = Field(description="페이지 ID"),
        webpart_id: str = Field(
            description=(
                "수정할 webPart ID (get_sharepoint_page의 webparts 리스트에서 받음). "
                "type이 'text'인 webPart만 안전하게 동작."
            )
        ),
        inner_html: str = Field(
            description=(
                "새 본문 HTML. 기존 innerHtml을 통째로 교체합니다. "
                "마크다운은 미지원이므로 사전 변환 필요. "
                "예: '<h2>2025년 매출</h2><p>130억 원</p>'"
            )
        ),
    ) -> dict[str, Any]:
        """SharePoint 페이지의 특정 webPart 본문(innerHtml)을 새 내용으로 교체합니다.

        ⚠️ 본문 교체는 사용자 검수 후 호출하세요. innerHtml은 통째로 새 내용으로
        덮어쓰이며, 기존 내용은 복구 불가합니다.

        정보 최신화 워크플로우:
        1. get_sharepoint_page로 webPart 목록과 현재 innerHtml 조회
        2. stale 정보 식별 (예: '2024년 매출 100억' → '2025년 매출 130억')
        3. 사용자와 새 inner_html 확정
        4. update_sharepoint_webpart 호출

        제한:
        - textWebPart(type='text')만 안전하게 지원. standardWebPart(이미지·링크·
          사람·문서 임베드 등)는 data.properties 스키마가 webPart마다 달라
          이 도구로 일괄 수정 불가. 해당 webPart는 SharePoint UI에서 직접 수정 권장.
        """
        graph = ctx.request_context.lifespan_context.graph
        body = {
            "@odata.type": "#microsoft.graph.textWebPart",
            "innerHtml": inner_html,
        }
        updated = await graph.patch(
            f"/sites/{site_id}/pages/{page_id}"
            f"/microsoft.graph.sitePage/webParts/{webpart_id}",
            json_body=body,
        )
        if updated is None:
            raise GraphAPIError("webPart 수정 응답이 비어있습니다")

        return {
            "webpart_id": updated.get("id"),
            "type": "text",
            "inner_html": updated.get("innerHtml"),
        }
