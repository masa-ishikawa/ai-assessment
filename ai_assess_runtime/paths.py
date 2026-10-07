"""生成器が使用するリポジトリ内パスの単一定義。"""

from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = PROJECT_DIR / "assessment_inputs"
CURRENT_INPUT_TEMPLATE = "ISV_AI_Use_Case_Assessment_Input_Template.xlsx"
OUTPUT_ROOT_DIR = PROJECT_DIR / "output"
PPTX_OUTPUT_DIR = OUTPUT_ROOT_DIR / "pptx"
JSON_OUTPUT_DIR = OUTPUT_ROOT_DIR / "json"
RESEARCH_OUTPUT_DIR = OUTPUT_ROOT_DIR / "research"

DEFAULT_ARCHITECTURE_IMAGE = PROJECT_DIR / "assets/architecture/oci_fixed_vm_adb_genai.png"
ORACLE_LOGO_IMAGE = PROJECT_DIR / "assets/architecture/Oraclelogo.png"
COVER_BACKGROUND_IMAGE = PROJECT_DIR / "assets/architecture/cover_background_ai_assessment.png"
ASSESSMENT_INTRO_IMAGE = PROJECT_DIR / "assets/ai_assess_intro.jpg"
PRIORITY_INSIGHT_IMAGE = PROJECT_DIR / "assets/icons/priority_insight.png"
