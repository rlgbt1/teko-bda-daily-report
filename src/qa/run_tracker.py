"""
qa/run_tracker.py — Build and export the final run-level audit record.

Produces:
  - FinalAudit  (Pydantic model, machine-readable)
  - JSON file   (saved to reports/audits/)
  - Markdown    (human-readable summary, also saved)

Usage:
    from src.qa.run_tracker import build_final_audit, save_audit

    audit = build_final_audit(
        run_id="2026-04-01",
        scrape_qa=scrape_qa_results,
        content_qa=content_qa_results,
        template_qa=template_result,
    )
    paths = save_audit(audit)
"""
from __future__ import annotations

import json
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from src.qa.schemas import (
    ContentQAResult,
    FinalAudit,
    QAStatus,
    ScrapeQAResult,
    TemplateQAResult,
)
from src.utils.logger import get_logger

logger = get_logger(__name__)

AUDIT_DIR = Path("reports/audits")


# ── Build ──────────────────────────────────────────────────────────────────────

def build_final_audit(
    scrape_qa:   list[ScrapeQAResult],
    content_qa:  list[ContentQAResult],
    template_qa: TemplateQAResult | None,
    run_id:      str = "",
) -> FinalAudit:
    """
    Aggregate QA results into a FinalAudit.

    Release gating logic:
    - safe_to_generate_ppt:   no FAIL-level scrape QA
    - safe_to_send_to_client: safe_to_generate_ppt AND template QA passes
    """
    run_id = run_id or datetime.utcnow().strftime("%Y%m%d-%H%M%S")

    # ── Derive statuses ───────────────────────────────────────────────────────
    scrape_statuses  = [r.status for r in scrape_qa]
    content_statuses = [r.status for r in content_qa]

    def _worst(statuses: list[QAStatus]) -> QAStatus:
        if QAStatus.FAIL in statuses:
            return QAStatus.FAIL
        if QAStatus.WARNING in statuses:
            return QAStatus.WARNING
        if statuses:
            return QAStatus.PASS
        return QAStatus.UNKNOWN

    scrape_worst  = _worst(scrape_statuses)
    content_worst = _worst(content_statuses)
    template_status = template_qa.status if template_qa else QAStatus.UNKNOWN

    all_statuses = [scrape_worst, content_worst, template_status]
    overall = _worst(all_statuses)

    # ── Confidence (average of non-zero scrape confidences) ───────────────────
    confidences = [r.confidence for r in scrape_qa if r.confidence > 0]
    overall_confidence = sum(confidences) / len(confidences) if confidences else 0.0

    # ── Release gates ─────────────────────────────────────────────────────────
    safe_to_generate_ppt = scrape_worst != QAStatus.FAIL
    safe_to_send = (
        safe_to_generate_ppt
        and content_worst != QAStatus.FAIL
        and (template_qa is None or template_qa.safe_to_release)
    )

    # ── Human-readable notes ──────────────────────────────────────────────────
    scrape_notes = []
    for r in scrape_qa:
        icon = {"pass": "✓", "warning": "⚠", "fail": "✗"}.get(r.status, "?")
        note = f"{icon} [{r.source}/{r.step}] {r.status}"
        if r.issues:
            note += f": {'; '.join(r.issues[:2])}"
        scrape_notes.append(note)

    content_notes = []
    for r in content_qa:
        icon = {"pass": "✓", "warning": "⚠", "fail": "✗"}.get(r.status, "?")
        note = f"{icon} [{r.section}] grounded={r.grounded}, safe={r.safe_to_include}"
        if r.issues:
            note += f": {'; '.join(r.issues[:2])}"
        content_notes.append(note)

    template_notes: list[str] = []
    slides_review:  list[str] = []
    if template_qa:
        template_notes = template_qa.deterministic_issues + template_qa.gemini_issues
        slides_review  = [f"Slide {i}" for i in template_qa.slides_needing_review]

    # ── Recommended action ────────────────────────────────────────────────────
    if overall == QAStatus.FAIL:
        action = "BLOQUEADO: corrigir erros antes de gerar o relatório."
    elif overall == QAStatus.WARNING:
        action = "AVISO: revisar manualmente os itens sinalizados antes de enviar ao cliente."
    else:
        action = "APROVADO: o relatório pode ser gerado e enviado."

    return FinalAudit(
        run_id=run_id,
        overall_status=overall,
        overall_confidence=round(overall_confidence, 3),
        safe_to_generate_ppt=safe_to_generate_ppt,
        safe_to_send_to_client=safe_to_send,
        scrape_qa=scrape_qa,
        content_qa=content_qa,
        template_qa=template_qa,
        scrape_integrity_notes=scrape_notes,
        content_grounding_notes=content_notes,
        template_compliance_notes=template_notes,
        slides_needing_review=slides_review,
        recommended_action=action,
    )


# ── Export ─────────────────────────────────────────────────────────────────────

def save_audit(audit: FinalAudit, output_dir: str | Path | None = None) -> dict[str, Path]:
    """
    Save the audit as JSON and Markdown.
    Returns {"json": Path, "markdown": Path}.
    """
    base = Path(output_dir) if output_dir else AUDIT_DIR
    base.mkdir(parents=True, exist_ok=True)

    slug = audit.run_id.replace(":", "-")
    json_path = base / f"audit_{slug}.json"
    md_path   = base / f"audit_{slug}.md"

    # JSON
    json_path.write_text(
        json.dumps(audit.model_dump(mode="json"), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    # Markdown
    md_path.write_text(_to_markdown(audit), encoding="utf-8")

    logger.info("Audit saved: %s | %s", json_path, md_path)
    return {"json": json_path, "markdown": md_path}


def _to_markdown(audit: FinalAudit) -> str:
    status_icon = {"pass": "✅", "warning": "⚠️", "fail": "❌"}.get(audit.overall_status, "❓")
    lines = [
        f"# BDA Daily Report — Audit {audit.run_id}",
        f"**Date:** {audit.timestamp.strftime('%Y-%m-%d %H:%M UTC')}",
        "",
        f"## Overall Status: {status_icon} `{audit.overall_status.upper()}`",
        f"- Confidence: {audit.overall_confidence:.0%}",
        f"- Safe to generate PPT: {'✅ Yes' if audit.safe_to_generate_ppt else '❌ No'}",
        f"- Safe to send to client: {'✅ Yes' if audit.safe_to_send_to_client else '❌ No'}",
        "",
        f"**Recommended action:** {audit.recommended_action}",
        "",
        "---",
        "",
        "## 1. Scrape Integrity",
    ]

    if audit.scrape_integrity_notes:
        for note in audit.scrape_integrity_notes:
            lines.append(f"- {note}")
    else:
        lines.append("- No scrape results recorded.")

    lines += [
        "",
        "## 2. Content Grounding",
    ]
    if audit.content_grounding_notes:
        for note in audit.content_grounding_notes:
            lines.append(f"- {note}")
    else:
        lines.append("- No content QA results recorded.")

    lines += [
        "",
        "## 3. Template Compliance",
    ]
    tq = audit.template_qa
    if tq:
        icon = {"pass": "✅", "warning": "⚠️", "fail": "❌"}.get(tq.status, "❓")
        lines.append(f"- Status: {icon} `{tq.status.upper()}`")
        lines.append(f"- Slide count: {tq.slide_count} (expected {tq.expected_slide_count})")
        lines.append(f"- Safe to release: {'✅ Yes' if tq.safe_to_release else '❌ No'}")
        if audit.template_compliance_notes:
            lines.append("- Issues:")
            for note in audit.template_compliance_notes:
                lines.append(f"  - {note}")
    else:
        lines.append("- Template QA not run.")

    if audit.slides_needing_review:
        lines += ["", "## 4. Slides Needing Manual Review"]
        for s in audit.slides_needing_review:
            lines.append(f"- {s}")

    lines += ["", "---", f"*Generated by teko-bda-daily-report QA layer*"]
    return "\n".join(lines) + "\n"
