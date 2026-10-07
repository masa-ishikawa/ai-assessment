"""通常生成パイプラインの明示的な依存・状態契約。

``generate_assessment.py`` は後方互換の公開ファサードであり、テストでは
その公開名を monkeypatch する。ここではその互換性を保ちながら、実行時に
必要な依存を一度だけ検証して固定し、生成段階の順序を型として表現する。
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from enum import Enum
from pathlib import Path
from typing import Any, Callable


class PipelineStage(str, Enum):
    """通常生成で許可する一方向の段階。"""

    INITIAL = "initial"
    INPUT_READY = "input_ready"
    STRATEGY_READY = "strategy_ready"
    ASSESSMENT_READY = "assessment_ready"
    RESEARCH_READY = "research_ready"
    PORTFOLIO_READY = "portfolio_ready"
    EVIDENCE_READY = "evidence_ready"
    QUANTITATIVE_READY = "quantitative_ready"
    DECISION_READY = "decision_ready"
    FROZEN = "frozen"
    RENDERED = "rendered"


_NEXT_STAGE: dict[PipelineStage, PipelineStage] = {
    PipelineStage.INITIAL: PipelineStage.INPUT_READY,
    PipelineStage.INPUT_READY: PipelineStage.STRATEGY_READY,
    PipelineStage.STRATEGY_READY: PipelineStage.ASSESSMENT_READY,
    PipelineStage.ASSESSMENT_READY: PipelineStage.RESEARCH_READY,
    PipelineStage.RESEARCH_READY: PipelineStage.PORTFOLIO_READY,
    PipelineStage.PORTFOLIO_READY: PipelineStage.EVIDENCE_READY,
    PipelineStage.EVIDENCE_READY: PipelineStage.QUANTITATIVE_READY,
    PipelineStage.QUANTITATIVE_READY: PipelineStage.DECISION_READY,
    PipelineStage.DECISION_READY: PipelineStage.FROZEN,
    PipelineStage.FROZEN: PipelineStage.RENDERED,
}


@dataclass(frozen=True, slots=True)
class PipelineState:
    """段階間で受け渡す生成状態。

    外側を frozen にして段階を飛び越えた代入を防ぐ。辞書は既存契約との
    互換性のため保持するが、各stageは新しい辞書を渡して ``advance`` する。
    """

    stage: PipelineStage = PipelineStage.INITIAL
    input_file: Path | None = None
    input_text: str = ""
    analysis_context: str = ""
    input_preprocessing: dict[str, Any] | None = None
    # 指定時だけ、外部JSONから正規化した固定15件カタログを保持する。原則として
    # PoC選定・詳細は今回の分析結果を使う通常経路で再作成するが、軽量カタログの
    # 明示poc_selectionだけはP1〜P3のID・順序を拘束できる。
    use_case_catalog: dict[str, Any] | None = None
    raw_company_name: str = ""
    resolved_company_name: str = ""
    midterm_plan: dict[str, Any] = field(default_factory=dict)
    assessment: dict[str, Any] = field(default_factory=dict)
    research: dict[str, Any] = field(default_factory=dict)
    front_matter_client: object | None = None
    research_client: object | None = None
    front_matter_sources: tuple[dict[str, Any], ...] = ()
    research_sources: tuple[dict[str, Any], ...] = ()
    # PORTFOLIO_READYで確定したP1〜P3の不変binding。後段はassessment内の
    # 途中生成物を再解釈せず、必ずこのスナップショットを参照する。
    final_poc_bindings: tuple[dict[str, str], ...] = ()
    executive_sources: tuple[dict[str, Any], ...] = ()
    executive_evidence: dict[str, Any] = field(default_factory=dict)
    research_path: Path | None = None
    review_payload: dict[str, Any] | None = None

    @property
    def source_text(self) -> str:
        """旧呼出し向けの読み取り専用alias。

        通常生成では ``input_text`` を原本、``analysis_context`` を正式社名や
        公開中計を付加した分析用文脈として扱う。旧テスト・外部参照が
        ``source_text`` を読む場合だけ、従来どおり分析用文脈を返す。
        """
        return self.analysis_context

    def advance(self, stage: PipelineStage, **changes: Any) -> "PipelineState":
        expected = _NEXT_STAGE.get(self.stage)
        if expected is not stage:
            raise RuntimeError(
                f"生成段階の遷移が不正です: {self.stage.value} -> {stage.value}"
                f"（期待: {expected.value if expected else '完了'}）"
            )
        # ``source_text=`` は旧内部APIとの互換入口として受け付ける。ただし、
        # 入力読込時だけ原本にも固定し、それ以降は分析用文脈だけを更新する。
        legacy_source_text = changes.pop("source_text", None)
        if legacy_source_text is not None:
            changes.setdefault("analysis_context", legacy_source_text)
            if self.stage is PipelineStage.INITIAL and stage is PipelineStage.INPUT_READY:
                changes.setdefault("input_text", legacy_source_text)
        if self.stage is not PipelineStage.INITIAL and "input_text" in changes:
            if changes["input_text"] != self.input_text:
                raise RuntimeError("入力原文input_textは読込完了後に変更できません。")
        return replace(self, stage=stage, **changes)


@dataclass(frozen=True, slots=True)
class GenerationServices:
    """通常生成が利用する依存を明示的に固定する。"""

    DEFAULT_ARCHITECTURE_IMAGE: Path
    DEFAULT_REASONING_EFFORT: str
    JSON_OUTPUT_DIR: Path
    PPTX_OUTPUT_DIR: Path
    QUANTITATIVE_EVIDENCE_SCHEMA_VERSION: str
    RESEARCH_OUTPUT_DIR: Path
    settings: Any
    analyze_assessment: Callable[..., dict]
    build_ai_product_business_impact: Callable[..., dict]
    build_consulting_front_matter: Callable[..., dict]
    build_executive_evidence_analysis: Callable[..., dict]
    build_final_poc_logic_details: Callable[..., list]
    build_midterm_plan_analysis: Callable[..., dict]
    build_poc_cost_estimate: Callable[..., dict]
    build_poc_decision_data: Callable[..., dict]
    build_review_payload: Callable[..., dict]
    collect_industry_research: Callable[..., list]
    collect_midterm_plan: Callable[..., dict]
    create_quantitative_analysis_client: Callable[..., object]
    create_oci_responses_client: Callable[..., object]
    create_pptx: Callable[..., None]
    derive_display_company_name: Callable[..., str]
    derive_display_service_name: Callable[..., str]
    extract_company_name: Callable[..., str]
    extract_isv_assessment_input: Callable[..., dict]
    extract_service_genre: Callable[..., str]
    fallback_consulting_front_matter: Callable[..., dict]
    apply_use_case_catalog: Callable[..., dict]
    load_isv_assessment_context: Callable[..., dict | None]
    load_use_case_catalog: Callable[..., dict]
    load_source_text: Callable[..., str]
    materialize_assessment_decision_contract: Callable[..., dict]
    materialize_ai_product_business_impact_contract: Callable[..., dict]
    materialize_quantitative_display_contract: Callable[..., dict]
    materialize_poc_portfolio: Callable[..., dict]
    normalize_adb_terminology: Callable[..., Any]
    normalize_isv_assessment_context: Callable[..., dict]
    normalized_use_case_label: Callable[..., str]
    primary_use_case_catalog_for: Callable[..., list[dict[str, str]]]
    canonical_use_case_id: Callable[..., str]
    resolve_input_file: Callable[..., Path]
    resolve_quantitative_display_contract: Callable[..., object]
    resolve_official_company_name: Callable[..., str]
    safe_filename: Callable[..., str]
    select_verified_quantitative_sources: Callable[..., dict]
    standard_document_page_count: Callable[..., int]
    validate_assessment_payload: Callable[..., list[str]]
    write_render_manifest: Callable[..., Path]

    @classmethod
    def from_api(cls, api: object) -> "GenerationServices":
        missing = [name for name in cls.__dataclass_fields__ if not hasattr(api, name)]
        if missing:
            raise RuntimeError("通常生成に必要な依存がありません: " + ", ".join(sorted(missing)))
        return cls(**{name: getattr(api, name) for name in cls.__dataclass_fields__})
