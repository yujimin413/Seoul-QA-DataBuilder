# 민관융합데이터 질의응답 챗봇 제작 프로젝트
## 서울 열린데이터 QA 학습 데이터 구축 및 정제 시스템

---
## 1. 개요
<img width="1000" alt="Image" src="https://github.com/user-attachments/assets/d5e209f6-bf6e-4fb0-a5a9-c3fe7d85802d" /><br>
서울 열린데이터 광장은 서울 생활인구, 수도권 생활이동, 서울 시민생활 데이터 등 다양한 빅데이터를 제공하고 있습니다.  
현재 댓글 기반 질의응답 방식으로 운영되고 있어 질문 증가에 따른 실시간 대응이 어렵고, 반복적인 질문이 많아 담당자의 업무 부담이 증가하고 있습니다.  
데이터 사용자의 편의성 향상을 위해 자동 응답 시스템이 필요한 상황입니다.

본 프로젝트는 이러한 문제를 해결하기 위해 **서울 생활인구, 수도권 생활이동, 서울 시민생활 데이터를 기반으로 GPTs 챗봇을 개발**하여 **반복적인 질의응답을 자동화**하고, **데이터 활용성을 높이며, 행정 업무 부담을 줄이는 것**을 목표로 합니다.  
이를 위해 이메일(.eml) 데이터를 분석하여 질문과 답변을 추출하고, 이를 주제별로 분류 및 정제하여 최종 학습 데이터셋을 구축합니다. 또한 OpenAI API를 활용하여 데이터를 정제하고 분석하며, 이를 기반으로 FAQ 추천 기능을 제공합니다.

---
## 2. 실행 순서

### 필수 단계

1. **1.extract_eml.py**: `.eml` 파일에서 질문과 답변을 추출하여 엑셀 파일로 저장
   - 현재 서울 열린데이터 광장에서는 시민들의 질문이 이메일로 전달되며, 이를 담당자가 수동으로 응답 후 웹사이트에 등록하는 방식으로 운영됨
![Image](https://github.com/user-attachments/assets/a8a71b9b-9e7f-45ad-9fd5-baaa519919be)

   - 질문-답변 추출 전
![Image](https://github.com/user-attachments/assets/3bd7a6f3-b8a5-43e5-b87a-fe1dde0aa0f5)

   - 질문-답변 추출 후
<img width="930" alt="Image" src="https://github.com/user-attachments/assets/74559b26-29d6-4e0c-a531-ff4239f5f095" />

2. **2.classify_qna_set.py**: OpenAI API를 활용하여 질문과 답변을 주제별로 분류
   - 질문의 내용에 따라 **생활인구, 생활이동데이터, 시민생활데이터** 등으로 분류하여 체계적인 응답을 제공
<img width="990" alt="Image" src="https://github.com/user-attachments/assets/b6746cb4-ac99-4c48-a692-039331447486" /><br><br>

3. **3.refine_qna_sets.py**: 질문과 답변을 정제하여 새로운 `문답모음집_정제.xlsx` 파일에 저장
   - OpenAI API를 활용해 질문과 답변을 정제하고, 중복된 질문을 통합하여 학습 데이터의 품질을 향상
   - 주제별 `.xlsx` 파일로 저장되며, 필요시 PDF로 변환하여 챗봇 학습 데이터로 활용 가능
<img width="990" alt="Image" src="https://github.com/user-attachments/assets/a98e2d1c-0a74-4b7c-a4dc-112ac63068d3" />
<img width="996" alt="Image" src="https://github.com/user-attachments/assets/cbdffd62-2572-494b-aed8-e3b45dbc240c" />

### 선택적 분석 및 평가 단계

4. **4-0.analyze_qna.py**: `문답모음집_정제.xlsx`을 분석하여 주요 질문 유형 및 빈도를 도출
   - 가장 많이 묻는 질문을 파악하여, 챗봇이 자주 묻는 질문에 신속하게 응답할 수 있도록 개선

5. **4-1.evaluate_consistency.py**: 분석 결과를 비교하여 질문 유형 일관성을 평가
   - 자카드 유사도, 코사인 유사도를 활용해 질문의 일관성을 분석하여 챗봇 응답의 정확성을 향상

6. **5.recommend_new_faq.py**: 기존 FAQ 데이터(`faq_dataset.xlsx`)와 비교하여 추가할 새로운 FAQ 추천
   - 기존 FAQ에 없는 질문을 자동으로 추천하여 챗봇의 학습 데이터를 지속적으로 업데이트 가능
  
<img width="1006" alt="Image" src="https://github.com/user-attachments/assets/0e88ebc7-5d13-4c83-a6ea-cfa4714b9fe2" />

---
## 3. 필수 파일

프로그램 실행을 위해 다음 파일이 필요합니다:

- `.env` (OpenAI API Key 및 제거할 개인정보 포함)
- `faq_dataset.xlsx` (기존 FAQ 데이터)
- `문답모음집.xlsx` (기본 질문-답변 데이터)
- `문답모음집_정제.xlsx` (정제된 최종 데이터, `3.refine_qna_sets.py` 실행 후 자동 업데이트됨)

### `.env` 파일 예시

`.env` 파일은 OpenAI API 키와 개인정보 보호를 위한 필터링 정보를 포함해야 합니다.

```env
# OpenAI API 키 (필수)
OPENAI_API_KEY=your-api-key

# 제거할 개인정보 목록 (예: 특정 이름, 이메일 주소)
REMOVE_NAMES=["홍길동", "김철수", "이영희"]
REMOVE_EMAILS=["user1@example.com", "user2@example.com"]
```

### `faq_dataset.xlsx` 파일 구조

- **시트명:**
  - 생활인구(문의)
  - 생활이동데이터(문의)
  - 시민생활데이터(문의)
- **컬럼:**
  - 질문
  - 답변

### `문답모음집.xlsx` 파일 구조

- **시트명:**
  - 생활인구(문의)
  - 생활이동데이터(문의)
  - 시민생활데이터(문의)
- **컬럼:**
  - 데이터셋명
  - 질문
  - 답변

### `문답모음집_정제.xlsx` 파일 구조

- **시트명:**
  - 생활인구(문의)
  - 생활이동데이터(문의)
  - 시민생활데이터(문의)
- **컬럼:**
  - 데이터셋명
  - 질문
  - 답변

---
## 4. 실행 방법

### 기본 실행

각 단계는 Python 환경에서 실행됩니다.

```bash
python 1.extract_eml.py
python 2.classify_qna_set.py
python 3.refine_qna_sets.py
```

### 추가 분석 및 평가 실행 (선택 사항)

```bash
python 4-0.analyze_qna.py
python 4-1.evaluate_consistency.py
python 5.recommend_new_faq.py
```

---
## 5. 프로젝트 폴더 구조

```
.
├── 1.extract_eml.py
├── 2.classify_qna_set.py
├── 3.refine_qna_sets.py
├── 4-0.analyze_qna.py
├── 4-1.evaluate_consistency.py
├── 5.recommend_new_faq.py
│
├── 1.DIRECTORIES_WITH_EML
│   └── sample/
│       ├── sample1.eml
│       ├── sample2.eml
├── 2.EXTRACTED_RESULTS
│   └── sample_추출.xlsx
├── 3.CLASSIFICATION_RESULTS
│   ├── sample_추출_생활인구(문의).xlsx
│   └── sample_추출_생활이동데이터(문의).xlsx
├── 4.REFINED_RESULTS
│   ├── sample_추출_생활인구(문의)_정제.xlsx
│   ├── sample_추출_생활이동데이터(문의)_정제.xlsx
│   ├── sample_추출_생활인구(문의)_정제.pdf
│   └── sample_추출_생활이동데이터(문의)_정제.pdf
├── QNA_ANALYZED_RESULTS
│   └── qna_analysis_results.xlsx
├── QNA_CONSISTENCY_RESULTS
│   └── consistency_analysis.xlsx
├── NEW_FAQ_RESULTS
│   └── new_faqs_results.xlsx
│
├── faq_dataset.xlsx
├── 문답모음집.xlsx
├── 문답모음집_정제.xlsx
├── .env
```

---
## 6. 참고사항

- 각 파일의 실행 방식과 세부 프로세스는 해당 `.py` 파일의 주석을 참고하세요.
- 선택적 단계(`4-0`, `4-1`, `5`)는 데이터 분석 및 일관성 검증을 위한 추가 기능입니다.
- 필수 단계(`1`, `2`, `3`)만 수행해도 `문답모음집_정제.xlsx`를 학습 데이터로 활용할 수 있습니다.
- OpenAI API를 활용한 데이터 정제 및 분석 기능이 포함되어 있으므로, `.env` 파일 내 API 키를 올바르게 설정해야 합니다.
- 챗봇은 **생활인구, 수도권 생활이동, 시민생활 데이터**에 대한 질문을 처리하며, PDF 문서 검색 기능을 활용하여 더욱 정확한 정보를 제공합니다.
- 최종적으로 챗봇은 **서울 열린데이터 광장**에 배포되며, 새로운 질문이 지속적으로 학습되어 챗봇 응답이 향상됩니다.
