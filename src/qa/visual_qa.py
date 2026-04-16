"""
qa/visual_qa.py — Visual layout QA for generated PPTX slides.

Pipeline:
  1. Export PPTX → PDF via AppleScript (PowerPoint must be installed)
  2. Render each PDF page → PNG via pymupdf
  3. Send (template_slide, generated_slide) image pair to GPT-4o-mini vision
  4. Return per-slide VisualSlideResult with issues and severity

Usage:
    from src.qa.visual_qa import VisualLayoutQA
    qa = VisualLayoutQA()
    results = qa.audit(generated_pptx_path, template_pptx_path)
    for r in results:
        print(r.slide_num, r.status, r.issues)
"""
from __future__ import annotations

import base64
import os
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from src.utils.logger import get_logger

logger = get_logger(__name__)

# Slides to audit — skip slide 2 (static agenda)
SLIDES_TO_AUDIT = [1, 3, 4, 5, 6, 7, 8, 9, 10, 11]

SLIDE_LABELS = {
    1:  "Cover",
    2:  "Agenda (static — skipped)",
    3:  "Sumário Executivo",
    4:  "Liquidez MN 1/2",
    5:  "Liquidez MN 2/2",
    6:  "Liquidez ME",
    7:  "Mercado Cambial",
    8:  "Mercado Capitais – BODIVA",
    9:  "Operações BDA",
    10: "Informação de Mercados 1/2",
    11: "Informação de Mercados 2/2",
}

_VISION_PROMPT = """You are auditing a PowerPoint slide for a Angolan bank (BDA) daily report.

Slide {slide_num}: {slide_label}

You are given TWO images:
  Image 1 = TEMPLATE (the reference — this is how it should look)
  Image 2 = GENERATED (the actual output — check this against the template)

Analyse the generated slide and report layout issues compared to the template.
Focus on:
- Text overflow or truncation (text cut off or spilling outside boxes)
- Shape/table overlaps (elements sitting on top of each other)
- Font size inconsistency (text noticeably larger or smaller than template)
- Missing sections or tables that exist in the template
- Wrong colours (headers, KPI ovals, section bars)
- Misaligned elements (tables/charts not aligned to template positions)
- Empty placeholder boxes that should contain data

Return ONLY valid JSON (no markdown fences) with this exact structure:
{{
  "status": "pass" | "warning" | "fail",
  "issues": ["concise description of each issue found"],
  "critical_overlaps": true | false,
  "font_ok": true | false,
  "missing_sections": ["section name if missing"],
  "summary": "1 sentence summary in English"
}}

If the slide looks correct compared to the template, return status "pass" with empty issues list.
Be concise — focus only on real, visible problems."""


@dataclass
class VisualSlideResult:
    slide_num: int
    slide_label: str
    status: str = "unknown"          # pass / warning / fail / skipped / error
    issues: list[str] = field(default_factory=list)
    critical_overlaps: bool = False
    font_ok: bool = True
    missing_sections: list[str] = field(default_factory=list)
    summary: str = ""
    llm_used: bool = False


class VisualLayoutQA:
    """
    Compares a generated PPTX against the template slide-by-slide using
    GPT-4o-mini vision. Falls back gracefully if PowerPoint or the API
    is unavailable.
    """

    def __init__(self) -> None:
        self._client = self._init_openai()

    def _init_openai(self):
        try:
            from openai import OpenAI
            api_key = os.getenv("OPENAI_API_KEY")
            if not api_key:
                logger.warning("OPENAI_API_KEY not set — visual QA disabled")
                return None
            return OpenAI(api_key=api_key)
        except ImportError:
            logger.warning("openai package not installed — visual QA disabled")
            return None

    @property
    def available(self) -> bool:
        return self._client is not None

    # ── Public API ─────────────────────────────────────────────────────────────

    def audit(
        self,
        generated_path: str,
        template_path: str,
        slides: list[int] | None = None,
    ) -> list[VisualSlideResult]:
        """
        Audit generated_path against template_path.

        Args:
            generated_path: Path to the generated .pptx
            template_path:  Path to the reference template .pptx
            slides:         1-based slide numbers to check. Defaults to SLIDES_TO_AUDIT.

        Returns:
            List of VisualSlideResult, one per audited slide.
        """
        slides_to_check = slides or SLIDES_TO_AUDIT

        if not self.available:
            return [
                VisualSlideResult(
                    slide_num=n,
                    slide_label=SLIDE_LABELS.get(n, f"Slide {n}"),
                    status="skipped",
                    summary="Visual QA skipped — OpenAI not configured",
                )
                for n in slides_to_check
            ]

        with tempfile.TemporaryDirectory() as tmpdir:
            logger.info("Exporting generated PPTX to PDF…")
            gen_pdf = self._export_to_pdf(generated_path, tmpdir, "generated.pdf")
            logger.info("Exporting template PPTX to PDF…")
            tmpl_pdf = self._export_to_pdf(template_path, tmpdir, "template.pdf")

            if not gen_pdf or not tmpl_pdf:
                logger.error("PDF export failed — cannot run visual QA")
                return [
                    VisualSlideResult(
                        slide_num=n,
                        slide_label=SLIDE_LABELS.get(n, f"Slide {n}"),
                        status="error",
                        summary="PDF export failed (is Microsoft PowerPoint installed?)",
                    )
                    for n in slides_to_check
                ]

            gen_images  = self._pdf_to_images(gen_pdf,  tmpdir, prefix="gen")
            tmpl_images = self._pdf_to_images(tmpl_pdf, tmpdir, prefix="tmpl")

        results = []
        for slide_num in slides_to_check:
            idx = slide_num - 1
            label = SLIDE_LABELS.get(slide_num, f"Slide {slide_num}")

            if idx >= len(gen_images) or idx >= len(tmpl_images):
                results.append(VisualSlideResult(
                    slide_num=slide_num,
                    slide_label=label,
                    status="error",
                    summary=f"Slide {slide_num} image not found after PDF export",
                ))
                continue

            logger.info("Auditing slide %d (%s)…", slide_num, label)
            result = self._audit_slide(
                slide_num, label, tmpl_images[idx], gen_images[idx]
            )
            results.append(result)

        return results

    # ── Rendering ─────────────────────────────────────────────────────────────

    def _export_to_pdf(self, pptx_path: str, out_dir: str, filename: str) -> str | None:
        """Export a PPTX to PDF using AppleScript + Microsoft PowerPoint."""
        out_path = str(Path(out_dir) / filename)
        abs_pptx = str(Path(pptx_path).resolve())
        abs_pdf  = str(Path(out_path).resolve())

        script = f'''
tell application "Microsoft PowerPoint"
    set pptxFile to POSIX file "{abs_pptx}"
    open pptxFile
    set theDoc to active presentation
    save theDoc in POSIX file "{abs_pdf}" as save as PDF
    close theDoc saving no
end tell
'''
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=60
            )
            if result.returncode == 0 and Path(abs_pdf).exists():
                logger.info("PDF export OK: %s", abs_pdf)
                return abs_pdf
            logger.error("AppleScript export failed: %s", result.stderr)
            return None
        except subprocess.TimeoutExpired:
            logger.error("AppleScript PDF export timed out")
            return None
        except Exception as exc:
            logger.error("PDF export error: %s", exc)
            return None

    def _pdf_to_images(
        self, pdf_path: str, out_dir: str, prefix: str, dpi_scale: float = 1.5
    ) -> list[str]:
        """Render each PDF page to a PNG. Returns list of file paths."""
        try:
            import fitz
        except ImportError:
            logger.error("pymupdf not installed — run: pip install pymupdf")
            return []

        doc = fitz.open(pdf_path)
        mat = fitz.Matrix(dpi_scale, dpi_scale)
        paths = []
        for i, page in enumerate(doc):
            out_path = str(Path(out_dir) / f"{prefix}_slide_{i+1:02d}.png")
            pix = page.get_pixmap(matrix=mat)
            pix.save(out_path)
            paths.append(out_path)
        doc.close()
        return paths

    # ── Vision call ───────────────────────────────────────────────────────────

    @staticmethod
    def _encode_image(path: str) -> str:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")

    def _audit_slide(
        self,
        slide_num: int,
        slide_label: str,
        template_img: str,
        generated_img: str,
    ) -> VisualSlideResult:
        prompt_text = _VISION_PROMPT.format(
            slide_num=slide_num,
            slide_label=slide_label,
        )

        tmpl_b64 = self._encode_image(template_img)
        gen_b64  = self._encode_image(generated_img)

        try:
            response = self._client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text",       "text": prompt_text},
                            {"type": "image_url",  "image_url": {"url": f"data:image/png;base64,{tmpl_b64}", "detail": "low"}},
                            {"type": "image_url",  "image_url": {"url": f"data:image/png;base64,{gen_b64}",  "detail": "low"}},
                        ],
                    }
                ],
                max_tokens=600,
                temperature=0.1,
            )
            raw_text = response.choices[0].message.content or "{}"
            # Strip markdown fences if present
            if raw_text.strip().startswith("```"):
                parts = raw_text.strip().split("```")
                raw_text = parts[1] if len(parts) > 1 else raw_text
                if raw_text.startswith("json"):
                    raw_text = raw_text[4:]

            import json
            data: dict[str, Any] = json.loads(raw_text.strip())

            return VisualSlideResult(
                slide_num=slide_num,
                slide_label=slide_label,
                status=data.get("status", "unknown"),
                issues=data.get("issues", []),
                critical_overlaps=bool(data.get("critical_overlaps", False)),
                font_ok=bool(data.get("font_ok", True)),
                missing_sections=data.get("missing_sections", []),
                summary=data.get("summary", ""),
                llm_used=True,
            )

        except Exception as exc:
            logger.error("Vision QA failed for slide %d: %s", slide_num, exc)
            return VisualSlideResult(
                slide_num=slide_num,
                slide_label=slide_label,
                status="error",
                summary=f"Vision call failed: {exc}",
            )
