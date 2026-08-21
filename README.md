# Consolidador de Nuevas OC (Cloud)

Streamlit Cloud에 배포해서 팀원들과 링크로 공유하는 버전입니다. 로컬 버전
(`oc_consolidator_clean/`)과 **완전히 동일한 엔진**을 사용합니다 — 고객사/제품
코드 자동 매칭(별칭 포함), 가격표 대조, 신용한도 체크까지 그대로 동작합니다.

## 로컬 버전과 다른 점은 딱 하나: 마스터 파일이 "미리 놓여있는 폴더"가 아니라 "매번 업로드"

로컬 버전은 PC의 특정 폴더에 미리 놓인 마스터 파일(customer master list, Product
master list, price list, credit report 등)을 자동으로 찾아서 씁니다. Streamlit
Cloud 서버는 회사 PC의 파일 시스템에 접근할 수 없기 때문에, 이 클라우드 버전은
**같은 마스터 파일들을 매번 브라우저로 업로드**받아서 씁니다. 업로드된 파일은
실행이 끝나면 서버에 남지 않고 삭제됩니다 (세션마다 임시 폴더를 새로 만들고,
처리 후 바로 지웁니다).

**"어제까지 누적된 결과 파일"도 업로드할 수 있습니다** — 로컬 버전처럼 새 OC를
그 파일 뒤에 이어붙이는 방식(중복 방지, 손으로 채운 컬럼 보존)이 그대로 동작합니다.

### ⚠️ 코드 유지보수 시 꼭 기억할 것

이 폴더(`oc_consolidator_cloud/`) 안의 `config.py`, `matcher.py`, `builder.py`,
`main.py`, `extractor.py`는 로컬 버전(`oc_consolidator_clean/`)의 **파일을 그대로
복사한 것**입니다 (Streamlit Cloud 배포는 하나의 폴더/저장소를 기준으로 하기 때문에,
자체 완결된 폴더로 만들어야 해서 이렇게 했습니다). **로컬 버전의 저 파일들을 나중에
수정하면, 이 클라우드 폴더에도 똑같이 복사해줘야 두 버전이 계속 같은 로직으로
동작합니다.** (Claude에게 "로컬이랑 클라우드 버전 엔진 파일 동기화해줘" 라고
요청하면 됩니다.)

## 1. 로컬에서 먼저 테스트하기

### 1-1. 필요한 프로그램 설치
Python 3.10 이상이 필요합니다. 설치되어 있는지 확인:
```bash
python --version
```

### 1-2. 이 폴더로 이동해서 라이브러리 설치
```bash
cd oc_consolidator_cloud
pip install -r requirements.txt
```

### 1-3. API 키 설정 (`.streamlit/secrets.toml`)
이 폴더에 이미 `.streamlit/secrets.toml` 파일이 만들어져 있습니다. 열어서
플레이스홀더 값을 실제 Anthropic API 키로 바꾸세요:

```toml
ANTHROPIC_API_KEY = "sk-ant-여기에-실제-키"
```

API 키는 https://console.anthropic.com 의 "API Keys" 메뉴에서 발급받습니다.

**이 파일은 절대 GitHub에 올리면 안 됩니다** — `.gitignore`에 이미 등록해뒀으니
`git add`를 해도 자동으로 제외됩니다. 실수로라도 커밋하지 않도록 주의하세요.

### 1-4. 실행
```bash
streamlit run app.py
```
브라우저가 자동으로 열립니다 (`http://localhost:8501`).

## 2. GitHub에 올리기

Streamlit Cloud는 GitHub 저장소를 기준으로 배포합니다.

1. GitHub에서 새 저장소를 만듭니다 (Private로 만드는 걸 권장 — 회사 업무용 코드이므로).
2. 이 `oc_consolidator_cloud` 폴더 안의 파일들을 그 저장소에 올립니다:
   ```bash
   cd oc_consolidator_cloud
   git init
   git add app.py requirements.txt .gitignore README.md
   git commit -m "Consolidador de Nuevas OC - version cloud"
   git branch -M main
   git remote add origin https://github.com/<본인계정>/<저장소이름>.git
   git push -u origin main
   ```
   (`.streamlit/secrets.toml`은 `.gitignore`에 있어서 자동으로 제외됩니다 — 확인차
   `git status`로 그 파일이 목록에 안 뜨는지 한 번 봐주세요.)

## 3. Streamlit Cloud에 배포하기

1. https://share.streamlit.io 접속 후 GitHub 계정으로 로그인
2. **"New app"** 클릭
3. 방금 만든 저장소, 브랜치(`main`), 메인 파일 경로(`app.py`)를 지정
4. **"Advanced settings" → "Secrets"** 에 아래 내용을 붙여넣기 (로컬 secrets.toml과 동일한 형식):
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-여기에-실제-키"
   ```
5. **"Deploy"** 클릭 — 몇 분 안에 빌드가 끝나고 `https://<앱이름>.streamlit.app` 같은
   공개 URL이 생성됩니다.

## 4. 팀원들과 공유할 때 꼭 확인할 것 (보안)

- 기본적으로 Streamlit Community Cloud 앱은 **URL을 아는 사람 누구나** 접속 가능합니다.
  회사 내부 구매주문 정보를 다루는 앱이므로, 다음 중 하나를 꼭 확인하세요:
  - Streamlit Cloud의 유료/팀 워크스페이스에서 "뷰어를 이메일로 제한" 옵션 사용
  - 또는 사내 SSO/VPN 뒤에 있는 별도 서버에 자체 호스팅
  - 최소한 URL을 회사 채널로만 공유하고, 외부에 노출되지 않도록 주의
- API 키는 Streamlit Cloud의 Secrets에만 저장되고, 앱을 쓰는 팀원들 화면에는 절대
  노출되지 않습니다 (코드에서 `st.secrets`로만 읽고 화면에 출력하지 않기 때문).
- 업로드된 파일은 세션이 끝나면 서버에 남지 않습니다 (메모리에서만 처리).

## 5. 파일 구성

```
oc_consolidator_cloud/
  app.py                    -> 웹 앱 UI 코드 (업로드 처리, 화면)
  config.py                 -> (로컬 버전 복사본) 마스터 파일/컬럼 위치 설정
  extractor.py               -> (로컬 버전 복사본) Claude로 OC 문서 읽기
  matcher.py                -> (로컬 버전 복사본) 마스터 파일 대조 로직
  builder.py                -> (로컬 버전 복사본) 엑셀 생성/업데이트 로직
  main.py                   -> (로컬 버전 복사본) 전체 흐름 오케스트레이션
  requirements.txt          -> 클라우드 배포 시 설치될 라이브러리 목록
  .streamlit/secrets.toml   -> 로컬 테스트용 API 키 (커밋 금지)
  .gitignore                -> secrets.toml이 실수로 커밋되지 않도록 제외
  README.md                 -> 이 파일
```
