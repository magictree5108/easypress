# 이지프레스 Easy Press

성동구 보도자료 초안·검수 도구 (개인 제작 프로토타입, Streamlit 버전)

기획안(HWP·HWPX·PDF·DOCX·TXT 업로드 또는 붙여넣기) → 팩트 추출(내부정보 자동 제외)
→ 제목 후보 5개 → 성동구 스타일 본문 생성 → 표기법·서술어 자동 검수 → 복사/다운로드

## 파일 구성

- `app.py` — 앱 전체
- `requirements.txt` — 의존성
- `.streamlit/config.toml` — 테마
- `assets/logo.png` — 로고 (직접 추가, 없어도 동작)

## 배포 (Streamlit Community Cloud, 무료)

1. GitHub에 새 저장소를 만들고 이 폴더의 파일을 전부 업로드한다
   (로고를 쓰려면 `assets/logo.png`로 추가)
2. https://share.streamlit.io 에 GitHub 계정으로 로그인
3. New app → 저장소·브랜치 선택, Main file path에 `app.py`
4. Advanced settings → Secrets에 아래를 입력

   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."   # console.anthropic.com에서 발급 (필수)
   PASSCODE = "0000"                  # 접근코드 (선택, 비우면 게이트 없음)
   MODEL = "claude-haiku-4-5"         # 선택, 기본값과 동일
   ```

5. Deploy → 생성된 `~.streamlit.app` URL을 공유하면 끝.
   받는 사람은 로그인 없이 브라우저에서 바로 사용한다.

## 로컬 실행

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."   # 또는 .streamlit/secrets.toml에 작성
streamlit run app.py
```

## 알아둘 것

- API 비용은 배포자 계정에서 나간다. console.anthropic.com에서
  월 지출 한도(예: $5)를 반드시 설정할 것. Haiku 기준 보도자료 1건에 수십 원 수준.
- 무료 호스팅은 한동안 접속이 없으면 잠들어 첫 접속이 수십 초 걸릴 수 있다.
- 입력 내용은 저장·로깅하지 않지만, 개인정보·비공개 문서는 넣지 않는 것이 원칙.
- HWP는 텍스트가 잘 추출되지만 표의 행·열 구조는 근사치다(셀 내용은 유지).
  PDF·DOCX 표는 행 단위(`셀 | 셀 | 셀`)로 구조가 보존된다.
- 스캔본 PDF(이미지)는 텍스트 추출이 안 되므로 붙여넣기를 안내한다.

## 문의

문의 및 피드백: 010-8829-5108(정호원)

본 도구는 성동구 공식 서비스가 아닌 개인 제작 프로토타입입니다.
