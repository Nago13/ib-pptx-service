"""
IB Presentation Microservice - FastAPI server que recebe JSON
e retorna apresentações PowerPoint formatadas para investment banking.
"""

import os
import logging
from fastapi import FastAPI, HTTPException
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

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


class SlideColumn(BaseModel):
    subtitle: str = ""
    bullets: list[str] = []


class TableData(BaseModel):
    headers: list[str] = []
    rows: list[list[str]] = []


class ChartSeries(BaseModel):
    name: str = ""
    values: list[float] = []


class ChartData(BaseModel):
    categories: list[str] = []
    series: list[ChartSeries] = []


class MetricItem(BaseModel):
    label: str = ""
    value: str = ""
    variation: str = ""


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


class PresentationRequest(BaseModel):
    presentation_title: str = "Apresentação"
    slides: list[SlideData]


@app.get("/health")
def health_check():
    return {
        "status": "healthy",
        "template_available": os.path.exists(TEMPLATE_PATH),
    }


@app.post("/generate")
def generate_presentation(request: PresentationRequest):
    try:
        logger.info(
            "Gerando apresentação '%s' com %d slides",
            request.presentation_title,
            len(request.slides),
        )

        template = TEMPLATE_PATH if os.path.exists(TEMPLATE_PATH) else None
        gen = IBPresentationGenerator(template_path=template)

        data = request.model_dump()
        pptx_bytes = gen.generate(data)

        safe_title = "".join(
            c if c.isalnum() or c in " -_" else "_"
            for c in request.presentation_title
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
def preview_slides(request: PresentationRequest):
    """Retorna a estrutura validada dos slides sem gerar o PPTX."""
    return {
        "title": request.presentation_title,
        "total_slides": len(request.slides),
        "layouts_used": [s.layout for s in request.slides],
        "slides": request.model_dump()["slides"],
    }
