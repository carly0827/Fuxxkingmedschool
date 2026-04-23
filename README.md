# 최종본 생성 사이트 (GitHub + Render 배포용)

이 저장소는 `줄.pdf + 야붙.pdf + 전사문`을 업로드하면 **최종본 PDF / safe_download ZIP / 검증 파일**을 생성하는 웹앱입니다.

## 포함 내용
- 기존 최종본 생성 파이프라인 코드 (`code/`)
- 웹 업로드 화면 (`app/`)
- Render 배포 설정 (`render.yaml`, `Dockerfile`)
- GitHub Actions 문법 체크 (`.github/workflows/python-check.yml`)

## 로컬 실행
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```
브라우저에서 `http://127.0.0.1:8000` 접속.

## GitHub에 올리는 방법
1. GitHub에서 새 repository 생성
2. 이 폴더 전체 업로드
3. 기본 브랜치를 `main`으로 두고 push

## 사이트로 배포하는 방법 (권장: Render)
1. Render 계정 생성
2. GitHub 연결
3. **New + → Web Service** 선택
4. 방금 올린 repo 선택
5. Render가 `render.yaml`을 읽도록 배포
6. 배포 완료 후 `https://...onrender.com` 주소 접속

## 주의
- GitHub Pages는 정적 사이트만 제공하므로 이 프로젝트처럼 Python/LibreOffice/PDF 처리가 필요한 앱은 실행할 수 없습니다.
- 생성 파일은 서버 저장소에 임시로 저장됩니다. 무료 인스턴스에서는 생성 직후 바로 다운로드하는 방식이 가장 안전합니다.

## 추천 전달 파일
생성 후에는 직접 PDF보다 `safe_download.zip`을 우선 받는 것을 권장합니다.
