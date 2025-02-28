# OpenAI API를 활용한 새로운 FAQ 추천 프로그램

# # 주요 기능:
# - 기존 FAQ 데이터와 새로운 Q&A 데이터를 비교하여 추가할 FAQ 추천
# - GPT를 이용하여 기존 FAQ에 없는 질문을 분석 및 추천 (터미널 출력)
# - 분석된 결과를 바탕으로 추가 FAQ 질문 추천 결과를 엑셀 파일로 저장
# - 필수 실행 프로그램은 아니며, FAQ 개선 시 활용 가능

# # 실행 조건:
# - OpenAI API 키 설정 필요 (`.env` 파일에 `OPENAI_API_KEY` 저장)
# - 기존 FAQ 데이터 (`faq_dataset.xlsx`) 필요
# - 분석할 Q&A 데이터 (`QNA_ANALYZED_RESULTS/` 폴더 또는 개별 파일) 필요

# # 입력 (Input):
# - `QNA_ANALYZED_RESULTS/` 폴더 내 저장된 분석 결과 (`.xlsx` 파일)
# - 또는 사용자가 직접 입력한 특정 `.xlsx` 파일 경로

# # 출력 (Output):
# - `NEW_FAQ_RESULTS/` 폴더에 분석된 결과 `.xlsx` 파일로 저장

import os
import openai
import pandas as pd
import matplotlib.pyplot as plt
import re
from collections import Counter
from dotenv import load_dotenv
from datetime import datetime
import matplotlib.pyplot as plt
import json

# 기본 디렉토리 설정
DEFAULT_INPUT_DIRECTORY = "QNA_ANALYZED_RESULTS"  # Q&A 분석 결과 폴더
DEFAULT_OUTPUT_DIRECTORY = "NEW_FAQ_RESULTS"  # FAQ 추천 결과 폴더
FAQ_FILE = "faq_dataset.xlsx"  # 기존 FAQ 데이터 파일
DEFAULT_OUTPUT_FILENAME = "new_faqs_results"  # 저장될 기본 파일명

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

# FAQ 추가 여부 판단 함수
def evaluate_faq_addition(sheet_name, qna_sets, faq_sets):
    """
    GPT를 활용하여 기존 FAQ와 비교 후, 추가할 새로운 질문을 선별
    """
    print(f"\n\n[INFO] '{sheet_name}' 시트의 신규 질문을 기존 FAQ와 비교 중...")

    # OpenAI API를 활용한 추가 여부 판단
    prompt = f"""

    # Persona
    당신은 데이터 분석 전문가이자 FAQ 관리 전문가입니다. 
    당신의 역할은 기존 FAQ와 새로운 질문-답변 데이터를 비교하여 **새로운 질문 유형을 식별**하고 **FAQ의 개선 여부를 평가**하는 것입니다.

    # Goals
    '{sheet_name}' 시트에서 수집된 기존 FAQ와 새로운 Q&A 데이터를 분석하여 다음과 같은 목표를 수행하세요:

    1. **기존 FAQ 평가**: 기존 FAQ가 사용자의 문의를 잘 반영하고 있는지 분석합니다.
    2. **새로운 질문 유형 식별**: 기존 FAQ에 없는 새로운 Q&A 질문-답변 쌍을 선별합니다.
    3. **추가할 FAQ 추천**: 기존 FAQ에서 누락된 중요한 질문을 찾아 추가할 질문-답변을 제안합니다.

    ---

    # Input Data
    ## 기존 FAQ 데이터
    {json.dumps(faq_sets, ensure_ascii=False, indent=2)}

    ## 새로운 Q&A 데이터
    {json.dumps(qna_sets, ensure_ascii=False, indent=2)}

    ---

    # Task
    ## 1. 기존 FAQ 평가
    - 기존 FAQ가 **사용자의 질문을 충분히 반영하고 있는지** 분석합니다.
    - 의미적으로 **중복된 항목이 있는지** 확인합니다.
    - 불필요하거나 구체성이 떨어지는 항목이 있다면 평가합니다.

    ## 2. 새로운 질문 유형 식별
    - 기존 FAQ에 포함되지 않은 **새로운 질문 유형**을 찾습니다.
    - 유사한 질문들을 **하나의 대표 질문**으로 병합하여 **일관성을 유지**합니다.
    - 의미가 동일한 질문들은 같은 유형으로 그룹화합니다.

    ## 3. 새로운 FAQ 추천
    - **기존 FAQ에 없는 질문 유형**을 기반으로 새롭게 추가해야 할 질문-답변 쌍을 추천합니다.
    - FAQ로 추가할 필요성이 높은 질문을 우선순위로 정리합니다.
    - 질문을 **사전순**으로 정렬하여 실행 간 변동성을 최소화합니다.

    ---

    # 출력 형식(JSON)
    {{
      "new_faqs": [
        {{
          "question": "추가할 질문 1",
          "answer": "추가할 답변 1",
          "priority": 1
        }},
        {{
          "question": "추가할 질문 2",
          "answer": "추가할 답변 2",
          "priority": 2
        }}
      ]
    }}

    ---

    # 주의사항:
    - **기존 FAQ와 비교하여 중복되지 않는 새로운 질문만 추천**하세요.
    - 질문 유형의 **일관성을 유지**하고, **표현을 통일**합니다.
    - 기존 FAQ에 없는 중요한 질문을 **우선순위로 추천**합니다.
    - 의미적으로 동일한 질문은 **병합하여 대표 질문**을 생성합니다.
    - 출력 결과는 **한국어**로 작성하며, 모든 질문은 **질문 형식**을 유지해야 합니다.

    """

    try:
        response = openai.ChatCompletion.create(
            model="gpt-4-turbo",
            messages=[
                {"role": "system", "content": "You are an expert in filtering and curating useful FAQs."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=2000,
            temperature=0.2,
            top_p=0.1
        )
        # GPT 응답에서 JSON 문자열 추출
        raw_response = response['choices'][0]['message']['content']

        print(f"## '{sheet_name}' 시트에 대한 GPT의 FAQ 분석 결과 ##")
        print(raw_response)


        # JSON 코드 블록만 추출
        extracted_json = extract_json_from_text(raw_response)

        if extracted_json:
            return json.loads(extracted_json)
        else:
            return {"new_faqs": []}

    except Exception as e:
        print(f"[ERROR] FAQ 추가 판단 중 오류 발생: {e}")
        return {"new_faqs": []}
    
def extract_json_from_text(response_text):
    """
    GPT 응답에서 JSON 코드 블록을 추출하는 함수
    """
    json_pattern = re.findall(r"```json\n(.*?)\n```", response_text, re.DOTALL)

    if json_pattern:
        return json_pattern[0]  # 첫 번째 JSON 블록 반환
    else:
        print(f"[ERROR] JSON 블록을 찾을 수 없습니다. 원본 응답: {response_text}")
        return None
    
# 기존 FAQ 데이터 불러오기
def load_existing_faq(faq_file):
    """
    기존 FAQ 데이터를 불러와 질문 목록을 set으로 저장
    """
    print("[INFO] 기존 FAQ 데이터 로드 중...")

    if not os.path.exists(faq_file):
        print(f"[WARNING] 기존 FAQ 데이터 파일 '{faq_file}'이 존재하지 않습니다.")
        return {}

    # 모든 시트 읽기
    df_sheets = pd.read_excel(faq_file, sheet_name=None)
    existing_faqs = {}

    for sheet_name, df in df_sheets.items():
        if '질문' not in df.columns or '답변' not in df.columns:
            print(f"[WARNING] 시트 '{sheet_name}'에 '질문' 또는 '답변' 컬럼이 없습니다. 스킵합니다.")
            continue

        # 기존 질문을 set으로 저장 (중복 방지)
        # existing_faqs[sheet_name] = set(df['질문'].dropna())
        existing_faqs[sheet_name] = list(zip(df['질문'].dropna(), df['답변'].dropna()))

    print("[INFO] 기존 FAQ 질문-답변 데이터 로드 완료.")
    return existing_faqs

def load_faq_candidates(analyzed_qna_file):
    """
    분석된 Q&A 데이터를 불러와 질문-답변 목록을 set으로 저장
    """
    print("[INFO] 분석된 Q&A 데이터 로드 중...")

    if not os.path.exists(analyzed_qna_file):
        print(f"[WARNING] 분석된 Q&A 데이터 파일 '{analyzed_qna_file}'이 존재하지 않습니다.")
        return {}

    # 모든 시트 읽기
    df_sheets = pd.read_excel(analyzed_qna_file, sheet_name=None)
    analyzed_qnas = {}

    for sheet_name, df in df_sheets.items():
        if '질문' not in df.columns or '답변' not in df.columns:
            print(f"[WARNING] 시트 '{sheet_name}'에 '질문' 또는 '답변' 컬럼이 없습니다. 스킵합니다.")
            continue

        # 기존 질문을 set으로 저장 (중복 방지)
        # existing_faqs[sheet_name] = set(df['질문'].dropna())
        analyzed_qnas[sheet_name] = list(zip(df['질문'].dropna(), df['답변'].dropna()))

    print("[INFO] 분석된 Q&A 질문-답변 데이터 로드 완료.")
    return analyzed_qnas
    
def main():
    print("=" * 80)
    print("OpenAI API를 활용한 새로운 FAQ 추천 프로그램")
    print("=" * 80)

    # 입력 방식 선택
    while True:
        choice = input(
            f"1. 기본 폴더({DEFAULT_INPUT_DIRECTORY})에서 파일 선택\n"
            "2. 직접 파일 경로 입력\n"
            "3. 프로그램 종료\n"
            "선택 (1, 2, 3): "
        ).strip()

        if choice == "1":
            available_files = [f for f in os.listdir(DEFAULT_INPUT_DIRECTORY) if f.endswith(".xlsx")]
            if not available_files:
                print(f"[WARNING] 기본 폴더({DEFAULT_INPUT_DIRECTORY})에 분석된 Q&A 파일이 없습니다. 다시 선택하세요.")
                continue

            print("\n기존 FAQ와 비교할 파일을 선택하세요 (번호 입력):")
            for idx, file_name in enumerate(available_files, start=1):
                print(f"{idx}. {file_name}")

            selected_index = input("\n파일 번호 입력: ").strip()
            if not selected_index.isdigit() or int(selected_index) < 1 or int(selected_index) > len(available_files):
                print("[ERROR] 올바른 번호를 입력하세요. 다시 시도하세요.")
                continue

            file_path = os.path.join(DEFAULT_INPUT_DIRECTORY, available_files[int(selected_index) - 1])
            break

        elif choice == "2":
            file_path = input("\n기존 FAQ와 비교할 파일의 경로를 입력하세요: ").strip()
            if not os.path.exists(file_path) or not file_path.endswith(".xlsx"):
                print("[ERROR] 유효한 .xlsx 파일 경로를 입력해야 합니다. 다시 시도하세요.")
                continue
            break

        elif choice == "3":
            print("프로그램을 종료합니다.")
            return

        else:
            print("[ERROR] 잘못된 입력입니다. 1, 2 또는 3을 입력하세요.\n")

    # 기존 FAQ 데이터 로드
    existing_faqs = load_existing_faq(FAQ_FILE)

    # 분석된 Q&A 데이터 로드
    analyzed_qnas = load_faq_candidates(file_path)

    # 실행 결과 저장할 딕셔너리
    all_new_faqs = {}

    # evaluate_faq_addition 실행 및 결과 저장
    for sheet_name in analyzed_qnas.keys():
        qna_sets = analyzed_qnas.get(sheet_name, [])
        faq_sets = existing_faqs.get(sheet_name, [])
        new_faqs_data = evaluate_faq_addition(sheet_name, qna_sets, faq_sets)

        print(f"## '{sheet_name}' 시트에 대한 GPT의 추천 질문 ##")
        print(new_faqs_data)

        # JSON 데이터를 pandas DataFrame으로 변환
        if "new_faqs" in new_faqs_data and new_faqs_data["new_faqs"]:
            df_new_faqs = pd.DataFrame(new_faqs_data["new_faqs"])
            
            # priority 순으로 정렬
            df_new_faqs = df_new_faqs.sort_values(by="priority", ascending=True)

            # 결과 저장
            all_new_faqs[sheet_name] = df_new_faqs

    # 저장할 파일명 (Unique Filename 적용)
    os.makedirs(DEFAULT_OUTPUT_DIRECTORY, exist_ok=True)
    output_file = get_unique_filename(DEFAULT_OUTPUT_DIRECTORY, DEFAULT_OUTPUT_FILENAME)

    # 새로운 FAQ를 Excel 파일로 저장
    if all_new_faqs:
        with pd.ExcelWriter(output_file) as writer:
            for sheet_name, df in all_new_faqs.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"[INFO] '{sheet_name}' 시트에 새로운 FAQ 추가 완료.")

        print(f"\n[INFO] 새로운 FAQ가 '{output_file}'에 저장되었습니다.")
    else:
        print("[INFO] 새로운 FAQ가 추가되지 않았습니다.")

    print("[INFO] 프로그램 실행 완료.")

# 실행
if __name__ == "__main__":
    main()

