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
    return {
        "status": "healthy",
        "template_available": os.path.exists(TEMPLATE_PATH),
        "generator_font": IBPresentationGenerator.FONT_TITLE,
    }


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
