# OpenAI API를 활용한 Q&A 데이터 분석 및 질문 유형 빈도 분석 프로그램

# 주요 기능:
# - OpenAI API를 이용하여 "문답모음집_정제.xlsx"에 저장된 질문 유형 분석 및 빈도 계산
# - 분석된 결과를 엑셀 파일로 저장
# - 시각화 가능 (현재는 주석처리 되어있음) 
# - 특정 데이터의 패턴을 분석할 때 유용하며 필수 실행 프로그램은 아님

# 실행 조건:
# - OpenAI API 키 설정 필요 (`.env` 파일에 `OPENAI_API_KEY` 저장)
# - 분석할 데이터 (`문답모음집_정제.xlsx`) 필요
# - 필수 파일: `문답모음집_정제.xlsx` (현재 디렉토리에 존재해야 함)
# - 시각화 할 경우(현재는 주석처리 되어있음) Malgun Gothic 글씨체 필요

# 입력 (Input):
# - `문답모음집_정제.xlsx` 파일 (질문 및 답변 포함)

# 출력 (Output):
# - `QNA_ANALYZED_RESULTS/` 폴더에 분석된 결과 `.xlsx` 파일로 저장

import os
import openai
import pandas as pd
import matplotlib.pyplot as plt
import re
from collections import Counter
from dotenv import load_dotenv
from datetime import datetime
import matplotlib.pyplot as plt

plt.rcParams['font.family'] ='Malgun Gothic'
plt.rcParams['axes.unicode_minus'] =False

# OpenAI API Key 설정
# .env 파일 로드 (파일이 없을 경우 종료)
if not os.path.exists(".env"):
    print("[WARNING] .env 파일이 존재하지 않습니다. 프로그램을 종료합니다.")
    exit()
else:
    load_dotenv()
openai.api_key = os.environ.get("OPENAI_API_KEY")

# 기본 디렉토리 설정
DEFAULT_OUTPUT_DIRECTORY = "QNA_ANALYZED_RESULTS"  # 분석 결과 저장 폴더
QNA_FILE = "문답모음집_정제.xlsx"  # 필수 입력 파일
QNA_ANALYSIS_FILE = "qna_analysis_results"  # 저장될 기본 파일명

def check_required_file():
    """
    필수 파일이 존재하는지 확인하는 함수
    """
    if not os.path.exists(QNA_FILE):
        print(f"[ERROR] 필수 파일 '{QNA_FILE}'이 존재하지 않습니다. 프로그램을 종료합니다.")
        exit()

# GPT를 활용한 질문 유형 빈도 분석 함수
def analyze_qna_patterns(sheet_name, qa_pairs):
    """
    OpenAI API를 활용하여 질문 유형을 분석하고, 유형별 빈도를 출력
    """
    print(f"[INFO] '{sheet_name}' 시트에 대한 질문-답변 패턴 분석을 시작합니다.")

    # 질문-답변 데이터를 텍스트로 변환
    qa_text = "\n".join([f"질문: {q}\n답변: {a}" for q, a in qa_pairs])

    prompt = f"""

    # Persona
    당신은 데이터 분석 전문가이며, 주어진 질문-답변 데이터를 분석하여 주요 질문 유형을 식별하는 역할을 합니다.

    # Goals
    아래는 '{sheet_name}' 시트에서 수집된 여러 **질문-답변 데이터**입니다.
    이 데이터를 분석하여 **2회 이상 등장하는 주요 질문 유형**을 선정하고, **각 유형의 출현 빈도를 계산**하세요.

    ## **질문-답변 데이터:**
    {qa_text}

    # Task
    ## 질문 유형 식별:
    - 동일한 의미의 질문들을 하나의 유형으로 묶습니다.
    - 의미적으로 유사한 질문을 강제적으로 하나로 병합하며, 실행 간 일관성을 유지합니다.
    - 질문 유형을 사전순 + 빈도순 정렬하여 실행 간 변동성을 최소화합니다.
    - 예를 들어, "데이터 조회 방법"과 "데이터를 어떻게 조회할 수 있나요?"는 같은 질문 유형으로 간주됩니다.

    ## 유사 질문 병합:
    - 의미적으로 유사한 질문을 고정된 대표 질문으로 변환합니다.
    - 대표 질문은 실제 사용자가 질문할 법한 자연스러운 문장 형식이어야 합니다.

    ## 빈도수 집계 및 정렬:
    - 같은 질문 유형으로 묶인 질문들의 빈도를 계산합니다.
    - 사전순 정렬 후 빈도순 정렬을 수행하여 실행 간 변동을 줄입니다.

    ## 주의사항: 
    - **질문 유형**이란, 동일한 유형의 질문을 그룹화한 것입니다. 예를 들어, "데이터 조회 방법", "통계 기준", "행정동 분석" 등의 패턴이 있습니다.
    - top_qna 출력은 출력 예시와 같은 형식으로하고, 한국어로 합니다.
    - top_qna top_qna_question은 질문 형식으로 합니다. 예를 들어, "2019년 생활이동 데이터의 제공 가능 여부"가 아니라 "2019년 생활이동 데이터 제공 가능한가요?"와 같은 형식으로 출력합니다.

    # 출력 형식:
    [(top_qna_question1, top_qna_answer1, frequency1), ..., (top_qna_questionN, top_qna_answerN, frequencyN)]

    ## 출력 예시:
    [
    (
        "서울시민생활데이터 데이터 분류 중 생활서비스 서비스 분류에서 배달 사용일수, 배달 브랜드 사용일수, 배달 식재료 사용일 수에 대해서 설명해주세요.",
        "배달 관련 서비스에 접속한 일수는 통신사가 관리하는 서비스 모두에 대해 일괄 합산된 것입니다. 예를 들어 배달의민족을 3일, 쿠팡이츠를 2일 사용했다면 총 5일 사용한 것으로 집계됩니다.",
        2
    ),
    (
        "서울시에서 향후 제공하는 생활인구 데이터의 갱신 주기는 어떻게 되나요?",
        "서울 생활인구는 일, 시각 단위로 생산되며, 갱신주기(측정주기)는 매일(5일전 데이터)입니다. 서울 생활인구는 열린데이터광장을 통해 일 단위로 갱신되어 제공됩니다.",
        10
    ),
    (
        "시민이 어떻게 활용할 수 있습니까?",
        "서울시 열린데이터광장(data.seoul.go.kr)을 통해 누구나 활용(행정동·1시간 단위)할 수 있으며, 서울시 빅데이터 캠퍼스에서는 보다 세분된 형태(250m 격자·20분 단위)의 데이터를 활용할 수 있습니다.",
        3
    ),
    (
        "'수도권 생활이동 데이터'를 개발한 이유는 무엇입니까?",
        "민선8기 정책기조에 따라 경기, 인천시민도 서울시민과 같은 정책수요자 임을 감안 수도권 도시민을 아우르는 정책개발을 위해 서울에서 수도권으로 확장된 수도권 생활이동 데이터를 개발하게 되었습니다. 수도권 생활이동 데이터는 다양한 정책에 활용이 가능합니다. 광역도시계획, 신도시 수요예측, 버스노선개선, 교통수요예측, 대중교통연계 그리고 지역상권활성화나 관광분야에도 두루 사용될 것으로 기대하고 있습니다.",
        4
    ),
    (
        "결측치 처리는 어떤 방식으로 진행하셨나요?",
        "통신데이터에서 결측치는 거의 없습니다. 특정 앱을 설치하지 않은 경우 관측값에는 결측치로 표시되지만, 실제 해당 앱을 사용하지 않은 경우로 0의 값을 가지며, 분석시에도 0 값으로 이용하였습니다. 데이터 오류 문제로 결측치가 있는 경우에는 값을 확인하고 분석에서 제외하였습니다.",
        6
    )
    ]

    """

    try:
        response = openai.ChatCompletion.create(
            model="gpt-4-turbo",
            messages=[
                {"role": "system", "content": "You are an expert in analyzing Q&A patterns."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=2000,
            temperature=0.2,
            top_p=0.01
        )

        # GPT 응답에서 질문 유형 및 빈도 추출
        patterns_text = response['choices'][0]['message']['content']

        # 정규식을 사용하여 (질문, 답변, 빈도) 쌍 추출
        matches = re.findall(r'\(\s*"(.*?)"\s*,\s*"(.*?)"\s*,\s*(\d+)\s*\)', patterns_text, re.DOTALL)

        # 질문 유형과 빈도 리스트 생성
        question_patterns = [(q.strip(), a.strip(), int(freq)) for q, a, freq in matches if q.strip() and a.strip() and freq.isdigit()]

        if not question_patterns:
            print(f"[WARNING] '{sheet_name}'에서 유효한 질문 유형을 찾지 못했습니다.")

        return question_patterns

    except Exception as e:
        print(f"[ERROR] 질문 유형 분석 중 오류 발생: {e}")
        return []

# 문답모음집을 분석하고 GPT 기반 빈도 분석 결과를 시각화하는 함수
def analyze_qna_file(qna_file):
    """
    문답모음집을 분석하여 GPT를 활용한 질문 유형 빈도 분석 및 시각화를 수행
    """
    print("[INFO] 문답모음집 분석을 시작합니다.")

    if not os.path.exists(qna_file):
        print(f"[ERROR] 파일 '{qna_file}'이 존재하지 않습니다.")
        return

    # 엑셀 파일 읽기 (모든 시트 로드)
    df_sheets = pd.read_excel(qna_file, sheet_name=None)

    # 결과 저장용 딕셔너리
    question_frequencies = {}

    for sheet_name, df in df_sheets.items():
        if '질문' not in df.columns or '답변' not in df.columns:
            print(f"[WARNING] 시트 '{sheet_name}'에 '질문' 또는 '답변' 컬럼이 없습니다. 스킵합니다.")
            continue

        print(f"[INFO] '{sheet_name}' 시트 분석 시작...")

        # 질문-답변 리스트 추출
        qa_pairs = list(zip(df['질문'].dropna(), df['답변'].dropna()))

        # GPT를 활용한 질문 유형 빈도 분석
        question_patterns = analyze_qna_patterns(sheet_name, qa_pairs)

        # 빈 결과도 저장하도록 수정
        question_frequencies[sheet_name] = question_patterns if question_patterns else []

        if question_patterns:
            question_frequencies[sheet_name] = question_patterns

            # 질문에 개행 추가
            max_line_length = 30  # 한 줄의 최대 길이
            def wrap_labels(label, max_length):
                return '\n'.join([label[i:i + max_length] for i in range(0, len(label), max_length)])
            
            labels, counts = zip(*[(wrap_labels(q, max_line_length), freq) for q, _, freq in question_patterns])
            
            # 그래프 그리기
            # plt.figure(figsize=(14, 8))  # 그래프 크기 조정
            # plt.barh(labels, counts, color='skyblue')
            # plt.xlabel("빈도")
            # plt.ylabel("질문 유형")
            # plt.title(f"{sheet_name} 시트의 질문 유형 빈도 분석")
            # plt.gca().invert_yaxis()  # y축 뒤집기
            # plt.subplots_adjust(left=0.4)  # 왼쪽 여백 추가
            # plt.show()

    if all(len(patterns) == 0 for patterns in question_frequencies.values()):
        print("[WARNING] 모든 시트에서 유효한 질문 유형이 없어 결과 파일을 생성하지 않습니다.")
        return

    return question_frequencies

def get_unique_filename(directory, base_filename):
    """
    동일한 파일명이 존재하면 번호를 추가하여 고유한 파일명 생성
    """
    filename = os.path.join(directory, f"{base_filename}.xlsx")
    counter = 1
    while os.path.exists(filename):
        filename = os.path.join(directory, f"{base_filename}({counter}).xlsx")
        counter += 1
    return filename

def save_analysis_results(question_frequencies):
    """
    분석된 질문 유형 빈도 데이터를 엑셀 파일로 저장
    """
    if not question_frequencies:
        print("[WARNING] 저장할 데이터가 없습니다. 파일 저장을 생략합니다.")
        return

    os.makedirs(DEFAULT_OUTPUT_DIRECTORY, exist_ok=True)

    output_file = get_unique_filename(DEFAULT_OUTPUT_DIRECTORY, QNA_ANALYSIS_FILE)

    with pd.ExcelWriter(output_file) as writer:
        for sheet_name, patterns in question_frequencies.items():
            df = pd.DataFrame(patterns, columns=["질문", "답변", "빈도"])
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"[INFO] '{sheet_name}' 시트에 분석 결과 저장 완료.")

    print(f"[INFO] 분석이 완료되었습니다! 결과는 '{output_file}' 파일을 확인하세요.")

def main():
    print("=" * 80)
    print("OpenAI API를 활용한 질문 유형 분석 프로그램")
    print("=" * 80)

    check_required_file()

    question_frequencies = analyze_qna_file(QNA_FILE)
    save_analysis_results(question_frequencies)

    print("[INFO] 프로그램 실행 완료.")

# 실행
if __name__ == "__main__":
    main()
