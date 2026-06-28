"""The CU Library — PowerPoint export API. Deploy on Railway."""
import re
from typing import Any, Dict, List
from fastapi import FastAPI
from fastapi.responses import Response, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from deck import build

app = FastAPI(title="The CU Library — PowerPoint Export")

# Allow the Vercel front-end (any origin) to call this from the browser.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["POST", "GET", "OPTIONS"],
    allow_headers=["*"],
)

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


class Deck(BaseModel):
    bookName: str = "Board Package"
    theme: Dict[str, Any] = {}
    period: Dict[str, Any] = {}
    sections: List[str] = []
    narrative: str = ""
    headline: List[Dict[str, Any]] = []
    balanceSheet: List[Dict[str, Any]] = []
    incomeStatement: List[Dict[str, Any]] = []
    spread: List[Dict[str, Any]] = []
    loanTrend: List[Dict[str, Any]] = []
    shareTrend: List[Dict[str, Any]] = []
    lending: Dict[str, Any] = {}
    lendingNarrative: str = ""
    ops: Dict[str, Any] = {}
    ceoNarrative: str = ""
    ceoHighlights: List[str] = []
    agenda: Dict[str, Any] = {}
    actionItems: List[Dict[str, Any]] = []
    securedSections: List[Dict[str, Any]] = []
    emerald: Dict[str, Any] = {}
    calendar: List[Any] = []
    execs: List[Any] = []
    acronyms: List[Any] = []


@app.get("/")
def health():
    return {"ok": True, "service": "cu-library-pptx"}


@app.post("/pptx")
def pptx(payload: Deck):
    try:
        data = build(payload.dict())
    except Exception as e:  # never 500 silently — surface a readable reason
        return JSONResponse(status_code=500, content={"error": str(e)})
    fname = (re.sub(r"[^\w]+", "-", payload.bookName or "board-package").strip("-") or "board-package") + ".pptx"
    return Response(content=data, media_type=PPTX_MIME,
                    headers={"Content-Disposition": 'attachment; filename="%s"' % fname})
