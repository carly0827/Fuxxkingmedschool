이 번들은 새로운 ChatGPT 세션에 업로드해 참고시키기 위한 패키지입니다.

구성:
- CHATGPT_workflow_guide.md : 작업 규칙/품질 기준/출력 규칙 설명
- paste_into_new_chatgpt_prompt.txt : 새 세션에 함께 붙여넣을 문구
- code/merge_jul_and_yabout.py : 줄+야붙 병합 코드
- code/transcript_note_generator.py : 전사문 정렬/필기본 생성/전체 전사문 정리본 DOCX→PDF 생성 코드
- code/build_final_notes_from_jul_yabout.py : 통합 실행 코드

이번 버전의 핵심:
1. 전체 전사문 정리본은 먼저 DOCX로 생성한 뒤 LibreOffice headless로 PDF로 변환합니다. 변환 실패 시에만 이미지 기반 PDF를 보조적으로 사용합니다.
2. 전체 전사문 정리본은 강의록 -> 시험용 정리표 관계를 모사해, 큰 섹션 단위로 묶고 슬라이드별 표 블록을 세로형으로 정리합니다.
3. 하단 오른쪽 요약 칸은 시간/전사 페이지 없이 `항목 | 정리` 표만 넣고, 구어체를 그대로 베끼지 않고 개념/대상/특징 중심으로 압축합니다.
4. 파란/빨간/보라 글씨 규칙은 유지됩니다.
5. 다운로드 오류를 줄이기 위해 번들 내부 파일명은 ASCII로 바꾸고 __pycache__를 제거했습니다.
6. 최종 결과물은 원래 한국어 zip 외에 핵심 결과만 담은 `*_safe_download.zip`, 전체 결과를 담은 `*_safe_full_archive.zip`도 함께 생성합니다.

권장 사용법:
1. 새로운 ChatGPT 대화에 이 번들을 업로드한다.
2. paste_into_new_chatgpt_prompt.txt 내용을 함께 보낸다.
3. 줄 파일, 야붙 파일, 전사문을 함께 제공한다.
4. 최종본 PDF, 전체 전사문 정리표 PDF, 검토 파일 생성을 요청한다.

다운로드/열기 문제 원인과 해결:
- 원인 1: 이전 번들에는 내부 zip 파일명이 비ASCII/깨진 이름으로 저장되어 iOS/일부 압축앱에서 열기 실패가 날 수 있었습니다.
- 원인 2: __pycache__와 불필요한 중간 파일이 들어 있어 용량이 커지고 앱 미리보기가 불안정할 수 있었습니다.
- 해결: 이번 번들은 내부 파일명을 ASCII로 통일했고, pycache를 제거했으며, 실행 결과에도 safe_download zip을 추가했습니다.

7. 전달 전 링크/열기 검증 단계가 추가되었습니다. 결과 폴더에 `*_delivery_link_check.json/txt`를 만들고, direct safe copy(ASCII 파일명) 기준으로 존재 여부/열림 가능 여부/PDF·ZIP 무결성을 확인합니다.
8. 사용자에게 링크를 줄 때는 한글/공백 원본 경로보다 direct safe copy 링크와 safe_download zip 링크를 우선 전달합니다.
