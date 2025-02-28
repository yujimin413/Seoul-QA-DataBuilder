# OpenAI API를 사용하여 질문-답변 데이터를 주제별로 분류하는 스크립트

# 주요 기능:
# - `.xlsx` 파일에서 질문과 답변을 읽어 생활인구, 생활이동데이터, 시민생활데이터로 분류
# - OpenAI API를 활용하여 질문과 답변 내용을 분석하고 적절한 주제로 분류
# - 분류된 데이터를 개별 파일로 저장하여 정리된 엑셀 파일로 출력

# 실행 조건:
# - Python 환경에서 실행 (`python 2.classify_qna_set.py`)
# - OpenAI API 키 설정 필요 (`.env` 파일에 `OPENAI_API_KEY` 저장)
# - `문답모음집.xlsx` 학습 파일이 실행 디렉토리에 존재해야 함
# - `1.eml_parsing.py` 실행 후 생성된 `{폴더명}_추출.xlsx` 파일이 `2.EXTRACTED_RESULTS/` 폴더에 있어야 함

# 입력 (Input):
# - `2.EXTRACTED_RESULTS/` 폴더 내 1.eml_parsing.py 실행 결과 (`{폴더명}_추출.xlsx`)
# - 엑셀 파일 컬럼: `질문`, `답변`

# 처리 방식:
# - 각 질문-답변을 OpenAI API를 활용하여 주제(생활인구, 생활이동데이터, 시민생활데이터)로 분류
# - 주제별로 데이터를 분리하여 저장

# 출력 (Output):
# - `3.CLASSIFICATION_RESULTS/` 폴더에 주제별 개별 파일로 저장
# - 파일명 형식: `{원본파일명}_{주제명}.xlsx`
# - 엑셀 파일 컬럼: `데이터셋명`, `질문`, `답변`
# - 각 주제(생활인구, 생활이동데이터, 시민생활데이터)에 대해 별도 파일 생성
# - 주제에 해당하는 데이터가 없으면 해당 파일은 생성되지 않음

# 사용 방법:
# 1. `1.eml_parsing.py` 실행 후 `2.EXTRACTED_RESULTS/` 폴더에 `{폴더명}_추출.xlsx` 파일이 생성되었는지 확인
# 2. 터미널에서 `python 2.classify_qna_set.py` 실행
# 3. `3.CLASSIFICATION_RESULTS/` 폴더에서 생성된 주제별 엑셀 파일 확인


import os
import openai
import pandas as pd
from dotenv import load_dotenv
import re

# 기본 디렉토리 설정
DEFAULT_INPUT_DIRECTORY = "2.EXTRACTED_RESULTS"  # 입력 데이터 폴더
DEFAULT_OUTPUT_DIRECTORY = "3.CLASSIFICATION_RESULTS"  # 결과 데이터 폴더
TRAINING_FILE = "문답모음집.xlsx"  # 학습용 데이터 파일

# OpenAI API Key 설정
# .env 파일 로드 (파일이 없을 경우 종료)
if not os.path.exists(".env"):
    print("[WARNING] .env 파일이 존재하지 않습니다. 프로그램을 종료합니다.")
    exit()
else:
    load_dotenv()
openai.api_key = os.environ.get("OPENAI_API_KEY")

# def classify_data(question, answer, model):
def classify_data(question, answer, training_file):
    """
    질문-데이터 set를 주제별로 분류하는 스크립트 제작

    """
    prompt = f"""
    # Task
    {training_file}을 기반으로 아래의 질문과 답변이 어떤 **주제**에 해당하는지 분류하고, 해당 **주제**에 맞는 시트로 분류해 주세요.
    질문: "{question}"
    답변: "{answer}"

    # 제약조건
    ## 출력 형식은 **반드시** 주제명(생활인구(문의) | 생활이동데이터(문의) | 시민생활데이터(문의))으로 지정해야 합니다.

    # **주제**는 다음과 같이 분류되어 있습니다.
    ## 주제1: 생활인구
    ## 주제2: 생활이동데이터
    ## 주제3: 생활이동

    # 저장 파일은 다음과 같은 시트로 구성되어 있습니다.
    ## Sheet1: "생활인구(문의)"    
    ## Sheet2: "생활이동데이터(문의)"
    ## Sheet3: "시민생활데이터(문의)"

    # 예시
    ## question, answer 예시 1
    "question": "생활이동 데이터에서 이동유형에 관하여 문의드립니다.\n\n서울시 생활이동 데이터에서 이동유형은 출발지와 도착지에 H(야간상주지), H(주간상주지), E(기타)를 부여하여 순서쌍으로 표현된 것으로 이해하고 있습니다.\n\n예를 들면, '11010'에서 '11020'으로의 이동유형이 HW이면, 출발지 '11010'은 해당 이동인구의 야간상주지(H)이고 도착지 '11020'는 주간상주지(W)가 된다고 생각합니다.\n\n그리고, 주야간상주지를 추정하는 방법에 대해서는 메뉴얼과 FAQ의 설명에 따르면 개인별로 주야간상주지가 정해지는 것으로 이해됩니다.\n\n그러면, 개인별로 주야간상주지가 하나 이상 부여받을 수 있는 건가요? 데이터에서 확인해보니 주야간상주지가 유일하게 부여되지 않은 것 같은 느낌을 받았습니다.\n\n예를 들면, 이동유형이 'HH'이면 주야간상주지가 같은 지역이라 출발지와 도착지가 같을 거라고 생각했지만 대부분 다른 지역이였고,\n\n반대로 출발지와 도착지가 같음에도 불구하고 이동유형 순서쌍이 다른 경우(HW, WH, 등)가 존재합니다.\n\n정리하자면, 주야간상주지는 개인별로 유일하게 추정되는지 궁금합니다.\n\n혹시 제가 이동유형이나 주야간상주지의 정의를 잘못 이해하고 있다면, 다시 알려주시면 감사하겠습니다.",
    "answer": "개인별로 주야간상주지는 한 달 동안의 체류 패턴에 따라 여러 개가 될 수 있습니다.\n\n체류지(야간/주간 상주지)는 일정 기간 이상 머문 장소를 기준으로 설정합니다."
    ## question, answer 분류 예시 1
    "생활이동데이터(문의)"

    ## question, answer 예시 2
    "question": "'서울생활인구 대도시권 내외국인' 데이터에서 서울 외 모든 지역의 거주자 수가 계산된 건가요? '생활인구순위' 열을 보면 대도시권 거주지가 68위까지 나와 있는데, 68위 밑으로는 생략된 건가요, 아니면 원래 68개밖에 없는 건가요?"
    "answer": "서울시 분석팀에서는 서울 외 모든 지역의 거주자 수 계산 여부에 대한 데이터를 명시하지 않았습니다. 또한, 12월 데이터 기준으로 '생활인구순위'는 68위까지 확인되며, 서울을 제외한 행정구역 코드는 68개로 확인되었습니다."
    ## question, answer 분류 예시 2
    "생활인구(문의)"

    ## question, answer 예시 3
    "question": "서울 시민생활 데이터 메뉴얼.pdf 의 12페이지를 보면, 초록색으로 된 기초통계량 표가 있습니다. 여기서 1인가구와 다인가구에 대한 기초통계량 정보가 기록되어 있는데, 1인가구에 대한 기초통계량은 주어진 데이터에는 없더라고요. 데이터 제작쪽에서만 가지고 있는 데이터인가요? 또한 다인가구의 평균 통화량을 실제로 구해봤는데, 값이 표와 일치하지 않더라고요. 표는 예시일 뿐이라서 그런건가요?"
    "answer": "안녕하세요. 매뉴얼 12페이지의 기초통계량은 데이터 제작 과정 중에 일회성으로 사용된 것으로서 현재는 폐기된 상태입니다. 해당 표의 내용은 예시로서 실제 값과 일치하지 않을 수 있습니다. 감사합니다."
    ## question, answer 분류 예시 3
    "시민생활데이터(문의)"

    ## Output Format
    생활인구(문의) | 생활이동데이터(문의) | 시민생활데이터(문의)

    """
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4-turbo",
            messages=[
                {"role": "system", "content": "You are an assistant specialized in refining and structuring data for AI training."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=1000
        )

        # API 응답에서 sheet명 추출
        sheet = response['choices'][0]['message']['content']

        # 정규표현식을 사용해 Sheet1, Sheet2, Sheet3 중 매치되는 문자열을 추출
        match = re.search(r"(생활인구\(문의\)|생활이동데이터\(문의\)|시민생활데이터\(문의\))", sheet)
        if match:
            sheet = match.group(0)

        ## sheet명 추출 및 반환
        # print("[INFO] sheet: ", sheet)
        return sheet

    except Exception as e:
        print(f"Error: {e}")
        return []
def get_unique_filename(directory, base_filename):
    """
    파일명이 중복될 경우 '(1)', '(2)' 등의 숫자를 추가하여 고유한 파일명을 반환.
    """
    filename = os.path.join(directory, f"{base_filename}.xlsx")
    counter = 1

    while os.path.exists(filename):
        filename = os.path.join(directory, f"{base_filename}({counter}).xlsx")
        counter += 1

    return filename


def process_file(input_file, training_file):
    """
    입력된 엑셀 파일을 처리하고 주제별로 개별 파일로 저장하는 함수.
    """
    df = pd.read_excel(input_file)
    if '질문' not in df.columns or '답변' not in df.columns:
        print(f"[ERROR] 파일 '{input_file}'에 '질문' 또는 '답변' 컬럼이 없습니다.")
        return

    classified_data = {
        "생활인구(문의)": [],
        "생활이동데이터(문의)": [],
        "시민생활데이터(문의)": []
    }

    dataset_names = {
        "생활인구(문의)": "생활인구문의",
        "생활이동데이터(문의)": "생활이동문의",
        "시민생활데이터(문의)": "시민생활문의"
    }

    base_filename = os.path.splitext(os.path.basename(input_file))[0]

    for index, row in df.iterrows():
        question = row['질문']
        answer = row['답변']
        # print(f"Processing row {index + 1}...")

        # API 호출
        sheet_name = classify_data(question, answer, training_file)

        # 각 주제별 데이터 저장
        classified_data[sheet_name].append({"질문": question, "답변": answer})

    # 결과 저장 폴더 생성
    os.makedirs(DEFAULT_OUTPUT_DIRECTORY, exist_ok=True)

    # 각 주제별 파일로 저장
    for sheet_name, data in classified_data.items():
        if data:
            output_base_filename = f"{base_filename}_{sheet_name}"
            output_file = get_unique_filename(DEFAULT_OUTPUT_DIRECTORY, output_base_filename)

            # 데이터프레임 생성 후 저장
            sheet_df = pd.DataFrame(data)
            sheet_df['데이터셋명'] = dataset_names[sheet_name]
            sheet_df[['데이터셋명', '질문', '답변']].to_excel(output_file, sheet_name=sheet_name, index=False)

            print(f"[INFO] '{output_file}'에 '{sheet_name}' 데이터를 저장했습니다.")
        if not data:
             print(f"[INFO] '{base_filename}' 파일에서 '{sheet_name}' 시트에 대한 데이터가 발견되지 않아 저장하지 않습니다.")

def main():
    print("=" * 80)
    print("STEP 2) OpenAI API를 사용하여 질문-답변 데이터를 주제별로 분류하는 프로그램")
    print("=" * 80)

    if not os.path.exists(TRAINING_FILE):
        print(f"[ERROR] 학습 파일 '{TRAINING_FILE}'을 찾을 수 없습니다. 프로그램을 종료합니다.")
        return

    while True:
        choice = input(f"1. 기본 폴더({DEFAULT_INPUT_DIRECTORY}) 내 모든 파일 처리\n"
                       "2. 특정 폴더 입력\n"
                       "3. 프로그램 종료\n"
                       "선택 (1, 2, 3): ").strip()

        if choice == "1":
            directory = DEFAULT_INPUT_DIRECTORY
            break
        elif choice == "2":
            directory = input("분류할 파일이 있는 폴더명을 입력하세요: ").strip()
            if not directory:
                print("[WARNING] 입력된 폴더명이 없습니다. 다시 입력하세요.")
                continue
            break
        elif choice == "3":
            print("프로그램을 종료합니다.")
            return
        else:
            print("[ERROR] 잘못된 입력입니다. 1, 2 또는 3을 입력하세요.\n")

    if not os.path.exists(directory):
        print(f"[ERROR] 디렉토리가 존재하지 않습니다: {directory}")
        return

    print(f"[INFO] '{directory}' 폴더 내 모든 .xlsx 파일을 처리합니다...")

    xlsx_files = [f for f in os.listdir(directory) if f.endswith(".xlsx")]

    if not xlsx_files:
        print(f"[WARNING] '{directory}' 내에 .xlsx 파일이 없습니다.")
        return

    for file_name in xlsx_files:
        file_path = os.path.join(directory, file_name)
        print(f"[INFO] '{file_name}' 처리 중...")
        process_file(file_path, TRAINING_FILE)

    print("\n[INFO] 모든 파일 처리가 완료되었습니다.")

# 실행 예제
if __name__ == "__main__":
    main()