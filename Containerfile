# OCI AI Assess PPTX generator. OCI認証情報はイメージに含めず、実行時にマウントします。
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN apt-get update \
    && apt-get install --no-install-recommends -y fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY generate_assessment.py assessment_input_workbook.py fill_assessment_input.py poc_cost_estimator.py poc_cost_config.py ./
COPY ai_assess_runtime ./ai_assess_runtime
COPY assets ./assets
COPY .agents/skills/ai-assessment ./.agents/skills/ai-assessment
COPY assessment_inputs/ISV_AI_Use_Case_Assessment_Input_Template.xlsx ./assessment_inputs/ISV_AI_Use_Case_Assessment_Input_Template.xlsx
RUN mkdir -p output/pptx output/json output/research

# 顧客入力と出力先はイメージへ含めず、/workへマウントして指定する。
# 例: -v "$PWD:/work" ai-assess /work/input.xlsx --output /work/output.pptx
VOLUME ["/work"]
ENTRYPOINT ["python", "/app/generate_assessment.py"]
CMD ["--help"]
