"""
IB Presentation Microservice - FastAPI server que recebe JSON
e retorna apresentações PowerPoint formatadas para investment banking.
"""

import os
import logging
from typing import Any

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, field_validator, model_validator

from generator import IBPresentationGenerator
from google_slides_generator import GoogleSlidesGenerator

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="IB Presentation Generator",
    description="Microserviço para gerar apresentações PowerPoint no padrão investment banking",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "templates")
TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "ib_master.pptx")
GOOGLE_CREDS_PATH = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "")


def _to_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v)


def _to_str_list(v: Any) -> list[str]:
    if not isinstance(v, list):
        return []
    return [str(item) if item is not None else "" for item in v]


def _to_float(v: Any) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


class SlideColumn(BaseModel):
    subtitle: str = ""
    bullets: list[str] = []

    @field_validator("subtitle", mode="before")
    @classmethod
    def coerce_subtitle(cls, v: Any) -> str:
        return _to_str(v)

    @field_validator("bullets", mode="before")
    @classmethod
    def coerce_bullets(cls, v: Any) -> list[str]:
        return _to_str_list(v)


class TableData(BaseModel):
    headers: list[str] = []
    rows: list[list[str]] = []

    @field_validator("headers", mode="before")
    @classmethod
    def coerce_headers(cls, v: Any) -> list[str]:
        return _to_str_list(v)

    @field_validator("rows", mode="before")
    @classmethod
    def coerce_rows(cls, v: Any) -> list[list[str]]:
        if not isinstance(v, list):
            return []
        return [_to_str_list(row) for row in v]


class ChartSeries(BaseModel):
    name: str = ""
    values: list[float] = []

    @field_validator("name", mode="before")
    @classmethod
    def coerce_name(cls, v: Any) -> str:
        return _to_str(v)

    @field_validator("values", mode="before")
    @classmethod
    def coerce_values(cls, v: Any) -> list[float]:
        if not isinstance(v, list):
            return []
        return [_to_float(item) for item in v]


class ChartData(BaseModel):
    categories: list[str] = []
    series: list[ChartSeries] = []

    @field_validator("categories", mode="before")
    @classmethod
    def coerce_categories(cls, v: Any) -> list[str]:
        return _to_str_list(v)


class MetricItem(BaseModel):
    label: str = ""
    value: str = ""
    variation: str = ""

    @field_validator("label", "value", "variation", mode="before")
    @classmethod
    def coerce_str(cls, v: Any) -> str:
        return _to_str(v)


class SlideData(BaseModel):
    layout: str
    title: str = ""
    subtitle: str = ""
    date: str = ""
    bullets: list[str] = []
    left_column: SlideColumn | None = None
    right_column: SlideColumn | None = None
    table: TableData | None = None
    chart_type: str = "bar"
    chart_data: ChartData | None = None
    metrics: list[MetricItem] = []
    contact_info: str = ""

    @field_validator("layout", "title", "subtitle", "date", "chart_type", "contact_info", mode="before")
    @classmethod
    def coerce_str(cls, v: Any) -> str:
        return _to_str(v)

    @field_validator("bullets", mode="before")
    @classmethod
    def coerce_bullets(cls, v: Any) -> list[str]:
        return _to_str_list(v)

    @model_validator(mode="before")
    @classmethod
    def handle_extra_fields(cls, data: Any) -> Any:
        if isinstance(data, dict):
            data.setdefault("layout", "content")
        return data


class PresentationRequest(BaseModel):
    presentation_title: str = "Apresentação"
    slides: list[SlideData]

    @field_validator("presentation_title", mode="before")
    @classmethod
    def coerce_title(cls, v: Any) -> str:
        return _to_str(v) or "Apresentação"


@app.get("/health")
def health_check():
    google_creds_ok = bool(
        GOOGLE_CREDS_PATH and os.path.exists(GOOGLE_CREDS_PATH)
    )
    return {
        "status": "healthy",
        "template_available": os.path.exists(TEMPLATE_PATH),
        "generator_font": IBPresentationGenerator.FONT_TITLE,
        "google_slides_available": google_creds_ok,
    }


@app.get("/diagnose")
def diagnose():
    """Tests Google API credentials and permissions for each service."""
    import json
    from google.oauth2 import service_account as sa
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError

    results: dict[str, Any] = {"credentials": {}, "tests": {}}

    if not GOOGLE_CREDS_PATH or not os.path.exists(GOOGLE_CREDS_PATH):
        results["credentials"]["error"] = (
            f"File not found: '{GOOGLE_CREDS_PATH}'"
        )
        return results

    try:
        with open(GOOGLE_CREDS_PATH) as f:
            creds_data = json.load(f)
        results["credentials"] = {
            "client_email": creds_data.get("client_email"),
            "project_id": creds_data.get("project_id"),
            "type": creds_data.get("type"),
            "file_path": GOOGLE_CREDS_PATH,
        }
    except Exception as e:
        results["credentials"]["error"] = str(e)
        return results

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/presentations",
        "https://www.googleapis.com/auth/drive",
    ]
    try:
        creds = sa.Credentials.from_service_account_file(
            GOOGLE_CREDS_PATH, scopes=scopes
        )
        results["credentials"]["scopes"] = scopes
    except Exception as e:
        results["credentials"]["auth_error"] = str(e)
        return results

    def _extract_error(exc: Exception) -> dict:
        info: dict[str, Any] = {"summary": str(exc)}
        if isinstance(exc, HttpError):
            info["status_code"] = exc.resp.status
            try:
                body = json.loads(exc.content.decode("utf-8"))
                info["error_body"] = body
                err = body.get("error", {})
                info["reason"] = err.get("status", "")
                details = err.get("details", [])
                for d in details:
                    if d.get("@type", "").endswith("ErrorInfo"):
                        info["error_reason"] = d.get("reason", "")
                        info["error_domain"] = d.get("domain", "")
                        info["error_metadata"] = d.get("metadata", {})
            except Exception:
                info["raw_content"] = exc.content.decode("utf-8", errors="replace")
        return info

    # --- Test Drive API ---
    try:
        drive = build("drive", "v3", credentials=creds)
        resp = drive.files().list(pageSize=1, fields="files(id,name)").execute()
        results["tests"]["drive"] = {
            "status": "OK",
            "files_found": len(resp.get("files", [])),
        }
    except Exception as e:
        results["tests"]["drive"] = {"status": "FAILED", **_extract_error(e)}

    # --- Test Sheets API ---
    try:
        sheets = build("sheets", "v4", credentials=creds)
        sheet = (
            sheets.spreadsheets()
            .create(body={"properties": {"title": "_diagnose_test_"}})
            .execute()
        )
        sid = sheet["spreadsheetId"]
        results["tests"]["sheets"] = {"status": "OK", "test_id": sid}
        try:
            drive = build("drive", "v3", credentials=creds)
            drive.files().delete(fileId=sid).execute()
            results["tests"]["sheets"]["cleanup"] = "deleted"
        except Exception:
            results["tests"]["sheets"]["cleanup"] = "could not delete"
    except Exception as e:
        results["tests"]["sheets"] = {"status": "FAILED", **_extract_error(e)}

    # --- Test Slides API ---
    try:
        slides_svc = build("slides", "v1", credentials=creds)
        pres = (
            slides_svc.presentations()
            .create(body={"title": "_diagnose_test_"})
            .execute()
        )
        pid = pres["presentationId"]
        results["tests"]["slides"] = {"status": "OK", "test_id": pid}
        try:
            drive = build("drive", "v3", credentials=creds)
            drive.files().delete(fileId=pid).execute()
            results["tests"]["slides"]["cleanup"] = "deleted"
        except Exception:
            results["tests"]["slides"]["cleanup"] = "could not delete"
    except Exception as e:
        results["tests"]["slides"] = {"status": "FAILED", **_extract_error(e)}

    return results


@app.post("/generate")
async def generate_presentation(request: Request):
    try:
        raw_body = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Body must be valid JSON")

    try:
        pres = PresentationRequest.model_validate(raw_body)
    except Exception as e:
        logger.warning("Validation failed, attempting auto-fix: %s", e)
        if isinstance(raw_body, dict):
            raw_body.setdefault("presentation_title", "Apresentação")
            raw_body.setdefault("slides", [])
        try:
            pres = PresentationRequest.model_validate(raw_body)
        except Exception as e2:
            logger.error("Validation still failed: %s", e2)
            raise HTTPException(status_code=400, detail=str(e2))

    try:
        logger.info(
            "Gerando apresentação '%s' com %d slides",
            pres.presentation_title,
            len(pres.slides),
        )

        template = TEMPLATE_PATH if os.path.exists(TEMPLATE_PATH) else None
        gen = IBPresentationGenerator(template_path=template)

        data = pres.model_dump()
        pptx_bytes = gen.generate(data)

        safe_title = "".join(
            c if c.isalnum() or c in " -_" else "_"
            for c in pres.presentation_title
        ).strip()
        filename = f"{safe_title or 'apresentacao'}.pptx"

        logger.info("Apresentação gerada com sucesso: %s", filename)

        return Response(
            content=pptx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    except ValueError as e:
        logger.error("Erro de validação: %s", e)
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.error("Erro ao gerar apresentação: %s", e, exc_info=True)
        raise HTTPException(status_code=500, detail=f"Erro interno: {str(e)}")


@app.post("/preview")
async def preview_slides(request: Request):
    """Retorna a estrutura validada dos slides sem gerar o PPTX."""
    try:
        raw_body = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Body must be valid JSON")

    pres = PresentationRequest.model_validate(raw_body)
    return {
        "title": pres.presentation_title,
        "total_slides": len(pres.slides),
        "layouts_used": [s.layout for s in pres.slides],
        "slides": pres.model_dump()["slides"],
    }


@app.post("/generate-slides")
async def generate_google_slides(request: Request):
    """
    Gera uma apresentação no Google Slides com gráficos vinculados
    a uma Google Sheet. Retorna as URLs para ambos os documentos.
    """
    try:
        raw_body = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Body must be valid JSON")

    try:
        pres = PresentationRequest.model_validate(raw_body)
    except Exception as e:
        logger.warning("Validation failed, attempting auto-fix: %s", e)
        if isinstance(raw_body, dict):
            raw_body.setdefault("presentation_title", "Apresentação")
            raw_body.setdefault("slides", [])
        try:
            pres = PresentationRequest.model_validate(raw_body)
        except Exception as e2:
            logger.error("Validation still failed: %s", e2)
            raise HTTPException(status_code=400, detail=str(e2))

    try:
        logger.info(
            "Gerando Google Slides '%s' com %d slides",
            pres.presentation_title,
            len(pres.slides),
        )

        gen = GoogleSlidesGenerator(credentials_path=GOOGLE_CREDS_PATH or None)
        data = pres.model_dump()
        result = gen.generate(data)

        logger.info(
            "Google Slides gerado: %s", result.get("slides_url", "")
        )

        return result

    except ValueError as e:
        logger.error("Erro de validação: %s", e)
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.error("Erro ao gerar Google Slides: %s", e, exc_info=True)
        raise HTTPException(status_code=500, detail=f"Erro interno: {str(e)}")
