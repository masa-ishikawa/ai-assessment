"""複数モジュール化した生成器全体の決定的フィンガープリント。"""

import hashlib
from pathlib import Path


ROOT_GENERATOR_FILES = (
    "generate_assessment.py",
    "assessment_input_workbook.py",
    "poc_cost_config.py",
    "poc_cost_estimator.py",
)

# PPTXの見た目・内容を決定する固定アセットも再現性契約へ含める。
# 画像を差し替えた場合に、Pythonコードが同じでも同一生成器とは扱わない。
ROOT_GENERATOR_ASSETS = (
    "assets/ai_assess_intro.jpg",
    "assets/architecture/Oraclelogo.png",
    "assets/architecture/cover_background_ai_assessment.png",
    "assets/architecture/oci_fixed_vm_adb_genai.png",
    "assets/icons/priority_insight.png",
)


def generator_source_files(project_dir: Path) -> tuple[Path, ...]:
    root_files = [project_dir / name for name in ROOT_GENERATOR_FILES]
    asset_files = [project_dir / name for name in ROOT_GENERATOR_ASSETS]
    runtime_dir = project_dir / "ai_assess_runtime"
    runtime_files = sorted(runtime_dir.rglob("*.py")) if runtime_dir.is_dir() else []
    return tuple(path for path in [*root_files, *runtime_files, *asset_files] if path.is_file())


def generator_source_sha256(project_dir: Path) -> str:
    """相対パスと内容を順序固定で集約し、生成コード全体を一つの値にする。"""
    digest = hashlib.sha256()
    for path in generator_source_files(project_dir):
        relative = path.relative_to(project_dir).as_posix().encode("utf-8")
        content = path.read_bytes()
        digest.update(len(relative).to_bytes(4, "big"))
        digest.update(relative)
        digest.update(len(content).to_bytes(8, "big"))
        digest.update(content)
    return digest.hexdigest()
