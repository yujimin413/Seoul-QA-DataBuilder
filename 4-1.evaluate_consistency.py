# OpenAI API를 활용한 Q&A 데이터 일관성 평가 프로그램

# 주요 기능:
# - 여러 개의 XLSX 파일을 입력받아 각 파일 내 시트별 "질문 유형"의 일관성을 평가
# - GPT를 활용하여 유사한 질문을 동일한 유형으로 분류
# - 분류된 질문에 대해 자카드 유사도와 코사인 유사도를 계산
# - 분석 결과를 바 차트로 시각화하고, 엑셀 파일로 저장
# - 특정 데이터의 일관성을 분석할 때 유용하며 필수 실행 프로그램은 아님

# 실행 조건:
# - OpenAI API 키 설정 필요 (`.env` 파일에 `OPENAI_API_KEY` 저장)
# - 분석할 데이터 (`QNA_ANALYZED_RESULTS/` 폴더 또는 개별 파일) 필요
# - 필수 폴더: `QNA_ANALYZED_RESULTS/` (자동 생성됨)
# - 3개 이상의 xlsx 파일 입력 추천

# 입력 (Input):
# - `QNA_ANALYZED_RESULTS/` 폴더 내 저장된 분석 결과 (`.xlsx` 파일)
# - 또는 사용자가 직접 입력한 특정 `.xlsx` 파일 경로

# 출력 (Output):
# - `QNA_CONSISTENCY_RESULTS/` 폴더에 분석된 결과 `.xlsx` 파일로 저장
# - 분석 결과를 바 차트로 시각화 가능 (현재 주석 처리)

import os
import openai
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from itertools import combinations
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from datetime import datetime
from dotenv import load_dotenv
import re

# 기본 디렉토리 설정
DEFAULT_INPUT_DIRECTORY = "QNA_ANALYZED_RESULTS"  # 기본 분석 결과 폴더
DEFAULT_OUTPUT_DIRECTORY = "QNA_CONSISTENCY_RESULTS"  # 결과 저장 폴더
DEFAULT_OUTPUT_FILENAME = "consistency_analysis"  # 저장될 기본 파일명

plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# OpenAI API Key 설정 (환경 변수에서 가져오기)
# .env 파일 로드 (파일이 없을 경우 종료)
if not os.path.exists(".env"):
    print("[WARNING] .env 파일이 존재하지 않습니다. 프로그램을 종료합니다.")
    exit()
else:
    load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

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

def collect_questions_from_sheets(file_paths):
    """
    여러 파일의 동일한 시트에서 질문들을 수집하여 한 리스트에 담는 함수.
    """
    sheets_data = [pd.ExcelFile(file_path).sheet_names for file_path in file_paths]
    common_sheets = set(sheets_data[0]).intersection(*sheets_data)

    questions_by_sheet = {}
    for sheet in common_sheets:
        print(f"\n[INFO] '{sheet}' 시트의 행 추출을 시작합니다.")
        questions = []
        for file_path in file_paths:
            print(f"[INFO] 파일: {file_path}에서 '{sheet}' 시트를 처리 중입니다.")
            df = pd.read_excel(file_path, sheet_name=sheet)
            if "질문" in df.columns:
                questions_extracted = df["질문"].dropna().tolist()
                print(f"[INFO] 추출된 질문 개수: {len(questions_extracted)}")
                questions.extend(questions_extracted)
        if questions:
            print(f"[INFO] '{sheet}' 시트에서 총 {len(questions)}개의 질문이 수집되었습니다.")
            questions_by_sheet[sheet] = questions
        else:
            print(f"[WARNING] '{sheet}' 시트에 질문이 없어 제외됩니다.")

    return questions_by_sheet

def classify_questions_with_gpt(sheet_name, questions):
    """
    GPT를 활용하여 시트명과 질문 리스트를 기반으로 유사한 질문을 분류합니다.
    """
    if not questions:
        print(f"[INFO] '{sheet_name}' 시트에 질문이 없어 분류를 건너뜁니다.")
        return []

    print(f"\n[INFO] '{sheet_name}' 시트의 질문 분류를 시작합니다.")

    prompt = f"""
    # Task
    다음은 "{sheet_name}" 시트에서 수집된 질문 리스트입니다. 
    의미적으로 유사하거나 동일한 질문끼리 하나의 튜플()에 담아 리스트[]로 반환하세요.

    ## 질문 리스트:
    {questions}

    ## Output Format
    1. 질문이 유사하거나 동일하다면 하나의 튜플로 묶어야 합니다. N은 자연수입니다. 예를 들어:
    [
        ("질문1", "질문2", "질문3"),
        ("질문4", "질문5"),
        ("질문6",),
        ...
        ("질문N-1", "질문N")
    ]
    2. **주의**: **단일 질문도 반드시 튜플 안에 넣어야 합니다.**  
        - 잘못된 예시: `"질문6"`  
        - 올바른 예시: `("질문6",)`
    3. 질문이 하나도 없거나 분류가 불가능한 경우 빈 리스트 []를 반환하세요.
    4. 응답에는 설명, 주석 또는 다른 텍스트가 포함되어서는 안 됩니다. 오직 리스트 형식의 데이터만 반환하세요.

    ## Output Example 1:
    [
        ("단기 체류 외국인의 정의가 무엇인가요?", "단기 체류 외국인의 정의는 무엇인가요?", "단기 체류 외국인의 정의는 무엇인가요?"),
        ("집계구 경계 shp 파일의 행정동 코드와 생활인구 파일의 행정동 코드가 다른 이유는 무엇인가요?",),
        ("서울 생활인구 데이터의 갱신 주기는 어떻게 되나요?", "서울 생활인구 데이터의 갱신 주기는 어떻게 되나요?", "서울 생활인구 데이터의 갱신 주기는 어떻게 되나요?"),
        ("생활인구 데이터에서 소수점이 나오는 이유는 무엇인가요?", "생활인구 데이터에서 소수점이 나오는 이유는 무엇인가요?")
    ]

    ## Output Example 2:
    [
        ("2019년 생활이동 데이터 제공 가능한가요?", "2019년 생활이동 데이터를 제공 받고 싶습니다."),
        ("서울시 실시간 도시데이터 API를 사용하여 공간적 범위를 설정하는 방법은 무엇인가요?", "서울시 실시간 도시데이터 API 사용 방법은 무엇인가요?"),
        ("교통폴리곤 단위의 데이터는 어디서 열람할 수 있나요?", "교통폴리곤 단위의 데이터는 어디서 확인할 수 있나요?"),
        ("생활이동 데이터에서 이동유형 HH와 WW의 의미는 무엇인가요?", "서울시 생활이동 데이터에서 HW, WH 이외의 이동 유형 HH, HE, WW, WE, EH, EW, EE는 각각 어떻게 정의되나요?", "이동유형 HH, WW의 정의와 자세한 설명을 알려주세요.")
    ]

    ## 주의사항
    1. 반드시 리스트([])와 튜플(()) 형식을 사용하세요. 하나의 원소를 가진 튜플도 () 안에 문자열을 입력해야 합니다.
    2. 불필요한 텍스트(설명, 주석, 분석 등)는 포함하지 마세요. 출력은 순수한 데이터만 포함해야 합니다.
    3. GPT가 올바르지 않은 형식으로 응답할 경우, 이를 방지하기 위해 출력 형식을 엄격히 준수하세요.

    """
    response = openai.ChatCompletion.create(
        model="gpt-4-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=2000,
        temperature=0.1
    )

    classified_text = response['choices'][0]['message']['content']
    
    print(classified_text)

    try:
        # 정규식으로 리스트 추출
        pattern = r'\[(.*?)\]'  # 리스트 구조를 찾는 정규식
        match = re.search(pattern, classified_text, re.DOTALL)
        if not match:
            print(f"[WARNING] '{sheet_name}' 시트에서 올바른 형식의 응답을 찾을 수 없습니다. -> 복구 시도")
            # extracted_text = classified_text.rstrip(",") + "]"  # 끊긴 경우 닫아줌
            fixed_text = re.sub(r"(,?\)\s*,\s*\))\s*$", ")]", classified_text)
            match = re.search(pattern, fixed_text, re.DOTALL)
            
        if match:
            extracted_text = match.group(0)
            classified_result = eval(extracted_text)  # 문자열로 반환된 리스트를 실제 리스트로 변환
            print(f"[INFO] '{sheet_name}' 시트의 질문 분류 결과: \n{classified_result}")
            return classified_result
        else:
            print(f"[WARNING] '{sheet_name}' 시트에서 올바른 형식의 응답을 찾을 수 없습니다.")
            return []
    
    except Exception as e:
        print(f"[ERROR] '{sheet_name}' 시트에서 질문 분류 중 오류 발생: {e}")
        return []
        
def calculate_similarity_within_tuple(sheet_name, grouped_questions):
    print(f"\n[INFO] '{sheet_name}' 시트의 질문 유사도 검사를 시작합니다.")
    """
    같은 튜플 내의 질문끼리 유사도를 계산하는 함수.
    - 단일 원소만 존재하는 튜플은 유사도 계산을 건너뜀
    - Jaccard Similarity & Cosine Similarity를 계산하여 결과 저장
    """
    jaccard_scores = []
    cosine_scores = []

    for question_group in grouped_questions:
        if not isinstance(question_group, (tuple, list)):  # 튜플 또는 리스트가 아닐 경우 예외 처리
            print(f"[ERROR] 잘못된 데이터 형식 감지: {question_group}. 유사도 계산을 건너뜁니다.")
            continue
        
        if len(question_group) == 1:
            print(f"[WARNING] '{question_group[0]}'는 비교 대상이 없어 유사도 검사를 스킵합니다.")
            continue  # 단일 질문이므로 건너뛰기

        print(f"[INFO] '{question_group[0]}' 그룹의 유사도 분석을 시작합니다.")

        try:
            # Ensure the input is a list of strings
            question_list = list(question_group)

            # 질문 벡터화 (TF-IDF)
            vectorizer = TfidfVectorizer()
            tfidf_matrix = vectorizer.fit_transform(question_list)
            cosine_matrix = cosine_similarity(tfidf_matrix)

            # 자카드 유사도 계산
            for i in range(len(question_list)):
                for j in range(i + 1, len(question_list)):
                    set1, set2 = set(question_list[i].split()), set(question_list[j].split())
                    jaccard_sim = len(set1 & set2) / len(set1 | set2) if len(set1 | set2) > 0 else 0
                    jaccard_scores.append(jaccard_sim)
                    cosine_scores.append(cosine_matrix[i, j])

            # 평균 유사도 계산
            avg_jaccard = np.mean(jaccard_scores) if jaccard_scores else 0
            avg_cosine = np.mean(cosine_scores) if cosine_scores else 0

            print(f"[INFO] '{question_group[0]}' 그룹의 유사도 분석 결과: "
                  f"Jaccard 평균 = {avg_jaccard:.4f}, Cosine 평균 = {avg_cosine:.4f}")

        except Exception as e:
            print(f"[ERROR] '{question_group[0]}' 그룹의 유사도 계산 중 오류 발생: {e}")

    # 전체 시트의 평균 유사도 계산
    final_avg_jaccard = np.mean(jaccard_scores) if jaccard_scores else 0
    final_avg_cosine = np.mean(cosine_scores) if cosine_scores else 0

    return final_avg_jaccard, final_avg_cosine

def visualize_results(consistency_scores):
    """
    유사도 분석 결과를 바 차트로 시각화합니다.
    """
    df = pd.DataFrame.from_dict(consistency_scores, orient="index")
    if not df.empty:
        df = df.sort_index()
        df.plot(kind='bar', figsize=(12, 6))
        plt.title("문의 유형별 유사도 분석")
        plt.xlabel("문의 유형")
        plt.ylabel("유사도")
        plt.xticks(rotation=0)
        plt.legend(loc='best')
        plt.show()
    else:
        print("[INFO] 시각화할 데이터가 없습니다.")

def save_results_to_excel(consistency_scores):
    """
    분석된 결과를 엑셀 파일로 저장
    """
    os.makedirs(DEFAULT_OUTPUT_DIRECTORY, exist_ok=True)
    output_file = get_unique_filename(DEFAULT_OUTPUT_DIRECTORY, DEFAULT_OUTPUT_FILENAME)

    df = pd.DataFrame.from_dict(consistency_scores, orient="index", columns=["Jaccard Similarity", "Cosine Similarity"])
    df.to_excel(output_file, index=True)
    print(f"[INFO] 분석 결과가 '{output_file}'에 저장되었습니다.")

def process_files(file_paths):
    """
    주어진 여러 개의 XLSX 파일에서 공통된 시트의 질문 데이터를 수집하고,
    GPT를 이용하여 유사한 질문을 그룹화한 후, 유사도 분석을 수행하는 함수.
    """
    print("-"*50)
    # 파일 내 공통된 시트에서 질문 데이터 수집
    questions_by_sheet = collect_questions_from_sheets(file_paths)

    # 결과 저장을 위한 딕셔너리
    consistency_scores = {}

    for sheet_name, questions in questions_by_sheet.items():

        print("-"*50)
        # GPT를 활용하여 질문 분류 수행
        grouped_questions = classify_questions_with_gpt(sheet_name, questions)
        # 같은 튜플 내에서만 유사도 계산
        avg_jaccard, avg_cosine = calculate_similarity_within_tuple(sheet_name, grouped_questions)
        print("-"*50)

        consistency_scores[sheet_name] = {
            'Jaccard Similarity': avg_jaccard,
            "Cosine Similarity": avg_cosine
        }

    return consistency_scores

def main():
    print("=" * 80)
    print("OpenAI API를 활용한 Q&A 데이터 일관성 평가 프로그램")
    print("=" * 80)

    while True:
        choice = input(
            f"1. 기본 폴더({DEFAULT_INPUT_DIRECTORY}) 내 파일 선택\n"
            "2. 비교할 파일 직접 경로 입력\n"
            "3. 프로그램 종료\n"
            "선택 (1, 2, 3): "
        ).strip()

        if choice == "1":
            # 기본 폴더 내의 모든 XLSX 파일 검색
            available_files = [f for f in os.listdir(DEFAULT_INPUT_DIRECTORY) if f.endswith(".xlsx")]

            if not available_files:
                print(f"[WARNING] 기본 폴더({DEFAULT_INPUT_DIRECTORY})에 비교할 .xlsx 파일이 없습니다. 다시 선택하세요.")
                continue

            print("\n[INFO] 기본 폴더 내에서 비교할 파일을 선택하세요 (쉼표로 구분하여 입력, 최소 2개 이상):")
            for idx, file_name in enumerate(available_files, start=1):
                print(f"{idx}. {file_name}")

            selected_indexes = input("\n파일 번호 입력 (쉼표로 구분, 최소 2개): ").strip().split(",")
            file_paths = [
                os.path.join(DEFAULT_INPUT_DIRECTORY, available_files[int(i) - 1])
                for i in selected_indexes if i.strip().isdigit() and 1 <= int(i) <= len(available_files)
            ]

            if len(file_paths) < 2:
                print("[ERROR] 최소 2개 이상의 파일을 선택해야 합니다. 다시 시도하세요.")
                continue
            break

        elif choice == "2":
            file_paths = input("\n비교할 파일의 경로를 쉼표로 구분하여 입력하세요 (최소 2개): ").strip().split(",")

            file_paths = [f.strip() for f in file_paths if os.path.exists(f.strip()) and f.strip().endswith(".xlsx")]

            if len(file_paths) < 2:
                print("[ERROR] 최소 2개 이상의 유효한 .xlsx 파일 경로를 입력해야 합니다. 다시 시도하세요.")
                continue
            break

        elif choice == "3":
            print("프로그램을 종료합니다.")
            return

        else:
            print("[ERROR] 잘못된 입력입니다. 1, 2 또는 3을 입력하세요.\n")

    # 파일 처리 및 유사도 분석 실행
    consistency_scores = process_files(file_paths)

    # 결과 시각화
    # visualize_results(consistency_scores)

    # # 결과 저장 (unique filename 적용)
    # os.makedirs(DEFAULT_INPUT_DIRECTORY, exist_ok=True)
    # base_filename = DEFAULT_OUTPUT_FILENAME
    # output_file = os.path.join(DEFAULT_INPUT_DIRECTORY, f"{base_filename}.xlsx")
    # counter = 1

    # while os.path.exists(output_file):
    #     output_file = os.path.join(DEFAULT_INPUT_DIRECTORY, f"{base_filename}({counter}).xlsx")
    #     counter += 1

    save_results_to_excel(consistency_scores)
    # print(f"[INFO] 분석 결과가 '{output_file}'에 저장되었습니다.")

if __name__ == "__main__":
    main()