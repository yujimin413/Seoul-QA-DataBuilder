# EML 파일에서 질문과 답변을 추출하여 엑셀 파일로 저장하는 스크립트

# 주요 기능:
# - `.eml` 파일에서 질문과 답변을 자동 추출하여 엑셀 파일로 저장
# - 불필요한 서명, 이전 이메일 본문, 개인정보(이메일, 이름 등) 제거
# - 추출된 데이터를 다음 단계(2.classify_qna_set.py)에서 활용 가능하도록 저장

# 실행 조건:
# - Python 환경에서 실행 (`python 1.extract_eml.py`)
# - `.env` 파일을 설정하여 개인정보 필터링할 목록을 JSON 형식으로 저장 가능
# - `1.DIRECTORIES_WITH_EML/` 폴더 내 `.eml` 파일들이 포함된 하위 폴더 필요

# 입력 (Input):
# - `1.DIRECTORIES_WITH_EML/` 폴더 내 `.eml` 파일이 포함된 여러 개의 하위 폴더
# - 각 `.eml` 파일에서 질문과 답변을 추출

# 출력 (Output):
# - `2.EXTRACTED_RESULTS/` 폴더에 `{폴더명}_추출.xlsx` 파일 생성
# - 엑셀 파일 컬럼: `질문`, `답변`

# 사용 방법:
# 1. `1.DIRECTORIES_WITH_EML/` 폴더에 `.eml` 파일들이 포함된 하위 폴더 준비
# 2. 터미널에서 `python 1.extract_eml.py` 실행
# 3. `2.EXTRACTED_RESULTS/` 폴더에서 생성된 엑셀 파일 확인

import email
import pandas as pd
import os
import re
from datetime import datetime
from dotenv import load_dotenv

import json

INPUT_DEFAULT_DIRECTORY = "1.DIRECTORIES_WITH_EML"
OUTPUT_DEFAULT_DIRECTORY = "2.EXTRACTED_RESULTS"

# 서명 제거를 위한 정규식 패턴 (환경변수에서 가져올 경우 JSON 형식으로 저장해야 함. 여기서는 직접 할당)
sig_patterns = [
    r"----------------------------------------------------------------------------------------------------------.*?----------------------------------------------------------------------------------------------------------",
    r"No.1 Digital value service company.*?유지관리팀"
]

# .env 파일 로드 (파일이 없을 경우 종료)
if not os.path.exists(".env"):
    print("[WARNING] .env 파일이 존재하지 않습니다. 프로그램을 종료합니다.")
    exit()
else:
    load_dotenv()

# .env 환경변수에서 JSON 형식의 리스트를 안전하게 변환
def get_env_list(env_var, default=[]):
    value = os.getenv(env_var, "[]")  # 기본값은 빈 리스트 문자열 "[]"
    try:
        return json.loads(value)  # JSON 문자열을 리스트로 변환
    except json.JSONDecodeError:
        print(f"[ERROR] 환경변수 {env_var} 값이 JSON 형식이 아닙니다. 기본값 사용.")
        return default  # 오류 발생 시 기본값 반환

# 환경변수에서 제거할 개인정보 리스트 가져오기
remove_names = get_env_list("REMOVE_NAMES")
remove_emails = get_env_list("REMOVE_EMAILS")

# 본문 추출
def parse_eml(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        msg = email.message_from_file(f)
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body += part.get_payload(decode=True).decode('utf-8', errors='ignore')
                # print("[DEBUG] 본문 (텍스트) :", body)
    else:
        body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
    return body

# 질문, 답변 추출
def extract_qna(body):
    # print(f"[DEBUG] 메일 본문\n{body}")

    # 서명 제거
    for pattern in sig_patterns:
        body = re.sub(pattern, "", body, flags=re.DOTALL)

    # print("\n[DEBUG] 발신자 및 수신자 서명 제거 후: \n", body)

    # 이전에 주고받은 메일의 본문을 질문에서 제외하기 위함
    # "보낸사람:" 또는 "From:" 두 번째 감지 후 내용 제거
    combined_pattern = r"(보낸사람:|From:)"
    matches = list(re.finditer(combined_pattern, body, flags=re.DOTALL))

    if len(matches) > 1:  # 두 번째 매칭이 있을 경우
        second_match_start = matches[1].start()  # 두 번째 매칭 시작 위치
        body = body[:second_match_start]  # 두 번째 매칭 이전 내용만 남김

    # print("\n[DEBUG] '보낸사람:' 또는 'From:' 두 번째 이후 제거 후: \n", body)

    # 불필요한 문자 제거
    body = re.sub(r"(안녕하세요.*|감사합니다.*|회신부탁드립니다.)+", "\n", body)
    body = re.sub(r".*답변드립니다[.!?]?", "", body)  # "답변드립니다"로 끝나는 문장 제거
    body = re.sub(r"From:.*|Sent:.*|To:.*|Cc:.*|Subject:.*", "\n", body)
    body = re.sub(r"보낸사람:.*|받는사람:.*|참조:.*|보낸시간:.*|제목:.*", "\n", body)
    body = re.sub(r"\n\s+", "\n", body)  # 여러 공백을 하나로 축소
    body = re.sub(r"(&nbsp;|\u200b|\xa0)+", "", body)  # HTML 공백 제거
    
    # print("\n[DEBUG] 불필요한 문자 제거 후: \n", body)
    # exit()

    # 본문을 "-----------------------원본 메세지-----------------------" 기준으로 나눔
    split_marker = "-----------------------원본 메세지-----------------------"
    if split_marker in body:
        parts = body.split(split_marker)
        answer = parts[0].strip()  # "답변" 부분
        # print("\n\n[DEBUG] 답변: \n", answer)
        question_block = parts[1] if len(parts) > 1 else ""
        # 질문 블록에서 메타데이터 제거
        question = re.sub(r"보낸사람:.*?제목:.*?\n+", "", question_block, flags=re.DOTALL).strip()
        # print("\n\n[DEBUG] 질문: \n", question)
    else:
        answer = body.strip()
        question = "정보 없음"

    # 개인정보 제거
    for name in remove_names:
        # if name in answer or name in question:
            # print(f"[DEBUG] '{name}' 발견! [###] 으로 대체")
        answer = answer.replace(name, "[###]")
        question = question.replace(name, "[###]")
    for email in remove_emails:
        # if email in answer or email in question:
        #     print(f"[DEBUG] '{name}' 발견! [###] 으로 대체")
        answer = answer.replace(email, "[###]")
        question = question.replace(email, "[###]")

    # print(f"##[DEBUG] 질문\n{question}\n\n##답변\n{answer}")
    # exit()

    return answer, question


def clean_answer(answer):
    # print(f"[DEBUG] 답변\n {answer}")
    # exit()

    return answer


def clean_question(question):
    # '○ 민원내용 :' 이후의 내용만 남기고 나머지 텍스트를 제거
    # 정규표현식을 사용하여 '○ 민원내용 :' 이후의 내용 추출
    match = re.search(r"○ 민원내용\s*:\s*\n?(.*?)\n○", question, re.DOTALL)
    if match:
        # '○ 민원내용 :' 이후부터 '○'로 시작하는 다음 섹션 전까지 내용 추출
        # print("[DEBUG] 패턴 매치 됨")
        question = match.group(1).strip()
        # print(f"[DEBUG] 추출 된 질문\n\n{extracted_content}")
        return question
    else:
        # 패턴이 없을 경우 원본 텍스트 반환
        return question

def process_eml_files(directory):
    records = []

    for file_name in os.listdir(directory):
        if file_name.endswith(".eml"):
            file_path = os.path.join(directory, file_name)

            body = parse_eml(file_path)
            answer, question = extract_qna(body)

            answer = clean_answer(answer)
            question = clean_question(question)
            
            records.append({
                "질문": question,
                "답변": answer,
            })

    # DataFrame 생성
    df = pd.DataFrame(records)
    return df

def save_parsed_data(df, directory):
    """
    결과 파일을 2.CLASSIFY_QNA_SET 폴더에 저장 (다음 프로그램에서 자동으로 사용할 수 있도록 함)
    """
    output_directory = OUTPUT_DEFAULT_DIRECTORY
    os.makedirs(output_directory, exist_ok=True)  # 폴더 생성

    base_output_file = f"{os.path.basename(directory)}_추출"
    output_file = os.path.join(output_directory, f"{base_output_file}.xlsx")
    counter = 1

    # 동일한 파일명이 존재하면 번호 붙이기
    while os.path.exists(output_file):
        output_file = os.path.join(output_directory, f"{base_output_file}({counter}).xlsx")
        counter += 1

    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"[INFO] 처리 완료. 결과는 '{output_file}'에 저장되었습니다.")


def main():
    print("=" * 80)
    print("STEP 1) EML 파일에서 질문과 답변을 추출하여 엑셀 파일로 저장하는 프로그램")
    print("=" * 80)

    # 기본 디렉토리 설정
    default_directory = INPUT_DEFAULT_DIRECTORY

    while True:
        choice = input(
            f"1. 기본 폴더({default_directory})의 모든 폴더 자동 처리\n"
            "2. 특정 폴더 입력\n"
            "3. 프로그램 종료\n"
            "선택 (1, 2, 3): "
        ).strip()

        if choice == "1":
            # 기본 폴더 내의 모든 하위 폴더 자동 검색
            directories = [
                os.path.join(default_directory, d)
                for d in os.listdir(default_directory)
                if os.path.isdir(os.path.join(default_directory, d))
            ]
            if not directories:
                print(f"[WARNING] 기본 폴더({default_directory})에 처리할 폴더가 없습니다.")
                continue
            break
        elif choice == "2":
            directories = input("분석할 폴더명을 쉼표(,)로 구분하여 입력하세요: ").strip().split(",")
            directories = [d.strip() for d in directories if d.strip()]
            if not directories:
                print("[WARNING] 유효한 폴더명이 입력되지 않았습니다. 다시 입력하세요.")
                continue
            break
        elif choice == "3":
            print("프로그램을 종료합니다.")
            exit()
        else:
            print("[ERROR] 잘못된 입력입니다. 1, 2 또는 3을 입력하세요.\n")

    for directory in directories:
        if not os.path.exists(directory):
            print(f"[WARNING] 디렉토리가 존재하지 않습니다: {directory}")
            continue

        print(f"[INFO] 디렉토리 처리 중: {directory}")
        df = process_eml_files(directory)
        save_parsed_data(df, directory)

if __name__ == "__main__":
    main()