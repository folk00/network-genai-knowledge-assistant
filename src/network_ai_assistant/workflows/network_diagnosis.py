from __future__ import annotations

from dataclasses import dataclass

from network_ai_assistant.rag.domain import RetrievalResult
from network_ai_assistant.rag.prompting import build_grounded_prompt


@dataclass(frozen=True)
class DiagnosisInput:
    incident_summary: str
    tenant_id: str = "demo-retail"
    site: str | None = None
    technology: str | None = None


def build_network_diagnosis_prompt(
    incident: DiagnosisInput,
    retrieved_context: list[RetrievalResult],
) -> str:
    """Controlled prompt for a network troubleshooting workflow."""

    question = (
        "Diagnose this network issue and provide evidence, likely cause, "
        "validation commands/checks, and safe next actions.\n\n"
        f"Incident: {incident.incident_summary}"
    )
    return build_grounded_prompt(question, retrieved_context)

