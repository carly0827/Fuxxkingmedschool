# 최종본 생성 사이트 (GitHub + Render 배포용)

이 저장소는 `줄 PDF + 야붙 PDF + 전사문 텍스트`를 업로드하면 최종본 PDF와 안전 다운로드 ZIP을 생성하는 웹앱입니다.

## 파일 구조
- `app.py`: 웹 서버 진입점
- `code/`: 사용자가 준 최종본 생성 파이프라인 코드
- `templates/index.html`: 업로드 화면
- `static/`: 스타일과 아이콘
- `Dockerfile`, `render.yaml`: Render 배포용 설정

## 로컬 실행
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --reload
```

## GitHub 업로드 방법
1. 이 zip을 먼저 압축 해제
2. GitHub repo에서 `Upload files`
3. 압축 해제된 폴더 안의 파일과 폴더를 한 번에 업로드
4. Commit changes

## Render 배포 방법
1. Render 계정 생성 후 GitHub 연결
2. New + → Web Service
3. 이 repo 선택
4. 배포 완료 후 `/healthz`가 `{"status":"ok"}`를 반환하는지 확인

## 주의
- GitHub Pages는 정적 페이지만 보여줄 수 있어서 실제 생성 기능은 실행되지 않습니다.
- 생성 직후 바로 결과 파일을 내려받는 것이 가장 안전합니다.
- 가장 먼저 받을 파일은 보통 `권장 다운로드 ZIP`입니다.
