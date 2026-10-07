"""標準デッキの論理スライド順序。

ページ数計算と描画順を同じ計画へ拘束し、片方だけを変更してページ番号が
ずれる回帰を防ぐ。
"""

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class SlideSpec:
    key: str
    title: str


def build_default_slide_plan(*, has_midterm: bool, use_case_page_keys: list[str],
                             include_cost_estimate: bool) -> tuple[SlideSpec, ...]:
    specs = [
        SlideSpec("cover", "表紙"),
        SlideSpec("intro", "OCI AI Use Case Assessmentとは"),
    ]
    if has_midterm:
        specs.append(SlideSpec("midterm", "中期経営計画とAI実装の整合性"))
    specs.extend((
        SlideSpec("industry_value", "業界・サービスの最新動向とAI活用"),
        SlideSpec("business_impact", "AI活用で期待できるビジネスインパクト"),
    ))
    specs.extend(SlideSpec(key, "AIユースケース一覧") for key in use_case_page_keys)
    specs.extend((
        SlideSpec("poc_selection", "15のAIユースケースから整理した AI技術アプローチの代表3テーマ"),
        SlideSpec("poc_detail:1", "AI技術テーマ 1"),
        SlideSpec("poc_detail:2", "AI技術テーマ 2"),
        SlideSpec("poc_detail:3", "AI技術テーマ 3"),
        SlideSpec("architecture", "代表3テーマを実現するOCI構成"),
    ))
    if include_cost_estimate:
        specs.append(SlideSpec("cost", "PoCにおけるOCI概算費用"))
    specs.extend((
        SlideSpec("support", "PoC実施・支援概要"),
        SlideSpec("closing", "Oracle"),
    ))
    return tuple(specs)
