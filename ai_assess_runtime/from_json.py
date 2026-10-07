"""承認済みレビューJSONの複製・移行・PPTX再描画経路。"""

import copy
import hashlib
import json
from pathlib import Path

from ai_assess_runtime.safe_io import atomic_render_immutable, atomic_write_json


_REPRODUCIBILITY_CONTRACT_V2_FORMAT = "ai-assess/reproducibility-v2"


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def run_from_json(args, parser, api) -> int:
    """外部AI・Web・再見積りを呼ばず、凍結済み契約だけを再生する。"""
    ASSESSMENT_JSON_FROZEN_FORMAT = api.ASSESSMENT_JSON_FROZEN_FORMAT
    DEFAULT_ARCHITECTURE_IMAGE = api.DEFAULT_ARCHITECTURE_IMAGE
    JSON_OUTPUT_DIR = api.JSON_OUTPUT_DIR
    PPTX_OUTPUT_DIR = api.PPTX_OUTPUT_DIR
    RESEARCH_OUTPUT_DIR = api.RESEARCH_OUTPUT_DIR
    build_poc_cost_estimate = api.build_poc_cost_estimate
    build_review_payload = api.build_review_payload
    cost_estimate_from_snapshot = api.cost_estimate_from_snapshot
    create_pptx = api.create_pptx
    load_assessment_payload = api.load_assessment_payload
    migrate_legacy_payload_for_review = api.migrate_legacy_payload_for_review
    safe_filename = api.safe_filename
    validate_assessment_payload = api.validate_assessment_payload
    write_render_manifest = api.write_render_manifest
    try:
        raw_payload = json.loads(args.from_json.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise ValueError(f"承認済みJSONを読み込めません: {error}") from error
    if not isinstance(raw_payload, dict):
        raise ValueError("承認済みJSONの最上位がオブジェクトではありません。")
    legacy_migrated = raw_payload.get("format") != ASSESSMENT_JSON_FROZEN_FORMAT
    reproducibility_migrated = False
    if legacy_migrated:
        if args.strict_json:
            raise ValueError(f"--strict-jsonでは{ASSESSMENT_JSON_FROZEN_FORMAT}以外を読み込めません。")
        migrated = migrate_legacy_payload_for_review(raw_payload)
        assessment_raw = migrated.get("assessment")
        research_raw = migrated.get("research")
        if not isinstance(assessment_raw, dict):
            raise ValueError("旧レビューJSONにassessmentオブジェクトがありません。")
        input_raw = raw_payload.get("input") if isinstance(raw_payload.get("input"), dict) else {}
        original_source_text = str(input_raw.get("source_text") or "")
        original_source_file = str(input_raw.get("source_file") or "").strip()
        payload = build_review_payload(
            assessment_raw,
            research_raw if isinstance(research_raw, dict) else {},
            source_file=Path(original_source_file) if original_source_file else None,
            source_text=original_source_text,
            preprocessing=input_raw.get("preprocessing"),
            cost_estimate=build_poc_cost_estimate(),
            include_source_text=args.include_source_text,
            architecture_image=args.architecture_image or DEFAULT_ARCHITECTURE_IMAGE,
        )
        errors = validate_assessment_payload(payload, strict=True)
        if errors:
            raise ValueError("旧レビューJSONの安全な移行に失敗しました: " + " / ".join(errors))
    else:
        # v2を再描画する場合はCLIの指定にかかわらずstrict契約を必須にする。
        # 「凍結形式だが、監査ログと表示値の対応が未確定」のJSONを
        # PowerPointだけに再生する逃げ道を残さない。
        # 明示的な再検証だけは、旧ルールでは凍結時に通過したが現行strict
        # 契約では失格となるclaimを、再凍結前に安全に除去できるようにする。
        # 構造は通常検証し、派生後のpayloadは必ずstrict検証する。
        payload = load_assessment_payload(
            args.from_json,
            strict=not getattr(args, "revalidate_midterm_targets", False),
        )
        provenance = payload.get("provenance") if isinstance(payload.get("provenance"), dict) else {}
        if args.json_only and provenance.get("format") != _REPRODUCIBILITY_CONTRACT_V2_FORMAT:
            # assessment-v2 / reproducibility-v1 は通常の再描画では従来どおり
            # そのまま再生する。一方、明示的な --json-only は移行操作として
            # 扱い、承認時の入力・費用を失わず現在のv2契約へ再凍結する。
            input_record = payload.get("input")
            rendering = payload.get("rendering") if isinstance(payload.get("rendering"), dict) else {}
            cost_snapshot = rendering.get("cost_estimate")
            if not isinstance(input_record, dict):
                raise ValueError("再現性契約v1の入力レコードをv2へ移行できません。")
            if cost_estimate_from_snapshot(cost_snapshot) is None:
                raise ValueError(
                    "再現性契約v1に有効な費用スナップショットがないため、"
                    "再見積りせずv2へ移行できません。"
                )
            render_profile = (
                rendering.get("render_profile")
                if isinstance(rendering.get("render_profile"), dict) else None
            )
            payload = build_review_payload(
                payload["assessment"],
                payload.get("research") if isinstance(payload.get("research"), dict) else {},
                source_file=None,
                source_text="",
                preprocessing=None,
                cost_estimate=cost_snapshot,
                include_source_text=False,
                architecture_image=args.architecture_image or DEFAULT_ARCHITECTURE_IMAGE,
                render_profile=render_profile,
                preserved_input_record=input_record,
            )
            errors = validate_assessment_payload(payload, strict=True)
            if errors:
                raise ValueError("再現性契約v1の安全なv2移行に失敗しました: " + " / ".join(errors))
            reproducibility_migrated = True
        if getattr(args, "revalidate_midterm_targets", False):
            # 凍結済みレビューを通常の再描画で変更しない原則は維持する。一方、
            # 明示的な再検証では、現行の決定的な原典照合ルールを適用してから
            # 新しいv2契約を作る。外部AI/Web/費用の再実行は行わない。
            input_record = payload.get("input")
            rendering = payload.get("rendering") if isinstance(payload.get("rendering"), dict) else {}
            cost_snapshot = rendering.get("cost_estimate")
            if not isinstance(input_record, dict):
                raise ValueError("凍結レビューJSONの入力レコードを再検証できません。")
            if cost_estimate_from_snapshot(cost_snapshot) is None:
                raise ValueError("有効な費用スナップショットがないため再検証できません。")
            render_profile = (
                rendering.get("render_profile")
                if isinstance(rendering.get("render_profile"), dict) else None
            )
            payload = build_review_payload(
                payload["assessment"],
                payload.get("research") if isinstance(payload.get("research"), dict) else {},
                source_file=None,
                source_text="",
                preprocessing=None,
                cost_estimate=cost_snapshot,
                include_source_text=False,
                architecture_image=args.architecture_image or DEFAULT_ARCHITECTURE_IMAGE,
                render_profile=render_profile,
                preserved_input_record=input_record,
            )
            errors = validate_assessment_payload(payload, strict=True)
            if errors:
                raise ValueError("中期経営計画の再検証に失敗しました: " + " / ".join(errors))
            reproducibility_migrated = True
    # v2は承認時のハッシュまで固定された凍結契約である。読み込み後に
    # 「決定的な補完」であっても内容を変えると、PPTXとmanifestが参照する
    # assessmentがずれる。旧JSONのみ上の移行分岐で補完し、v2へ再凍結・
    # strict再検証してからここへ到達する。
    assessment = copy.deepcopy(payload["assessment"])
    research = copy.deepcopy(payload.get("research", {}))
    rendering = payload.get("rendering") if isinstance(payload.get("rendering"), dict) else {}
    cost_estimate = cost_estimate_from_snapshot(rendering.get("cost_estimate"))
    if args.json_only:
        if args.output:
            parser.error("--json-onlyでは--outputではなく--output-jsonを指定してください。")
        json_path = args.output_json or JSON_OUTPUT_DIR / (
            f"{safe_filename(str(assessment['service_name']))}_assessment.json"
        )
        json_path.parent.mkdir(parents=True, exist_ok=True)
        atomic_write_json(json_path, payload)
        RESEARCH_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        research_path = RESEARCH_OUTPUT_DIR / f"{safe_filename(str(assessment['service_name']))}_research.json"
        atomic_write_json(research_path, research)
        action = "安全なv2レビューJSONへ移行・再検証" if (legacy_migrated or reproducibility_migrated) else "複製"
        print(f"承認済みJSONを{action}しました: {json_path}")
        print(f"      リサーチログ: {research_path}")
        return 0
    architecture_image = args.architecture_image or DEFAULT_ARCHITECTURE_IMAGE
    if not architecture_image.is_file():
        parser.error(f"構成図画像が見つかりません: {architecture_image}")
    provenance = payload.get("provenance") if isinstance(payload.get("provenance"), dict) else {}
    if provenance.get("format") == _REPRODUCIBILITY_CONTRACT_V2_FORMAT:
        architecture_asset = (
            rendering.get("architecture_asset")
            if isinstance(rendering.get("architecture_asset"), dict) else {}
        )
        frozen_sha256 = str(architecture_asset.get("sha256") or "")
        if _file_sha256(architecture_image) != frozen_sha256:
            raise ValueError(
                "指定された構成図画像が凍結レビューJSONのSHA-256と一致しません。"
                "承認時と同じ画像を指定してください。"
            )
    output_path = args.output or PPTX_OUTPUT_DIR / f"{safe_filename(str(assessment['service_name']))}_AIユースケース.pptx"
    if output_path.suffix.lower() != ".pptx":
        parser.error("--outputは.pptxファイルを指定してください。")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if cost_estimate is None:
        raise ValueError("凍結レビューJSONに有効な費用スナップショットがないため、PPTXを描画できません。旧JSONは--from-json --json-onlyでv2へ移行してください。")
    atomic_render_immutable(
        output_path,
        assessment,
        lambda temporary_output: create_pptx(
            assessment, architecture_image, temporary_output, cost_estimate, research,
        ),
        mutation_message=(
            "PPTX描画処理が承認済みassessmentを変更しました。"
            "凍結JSONの再現性を保てないため出力を中止します。"
        ),
    )
    manifest_path = write_render_manifest(
        output_path, payload, artifact_format="pptx", design="default",
    )
    print(f"承認済みJSONからPPTXを出力しました: {output_path}")
    print(f"      再現性manifest: {manifest_path}")
    return 0
