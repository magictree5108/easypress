# -*- coding: utf-8 -*-
"""
이지프레스 Easy Press — 성동구 보도자료 초안·검수 도구 (Streamlit)
로그인 없이 URL로 쓰는 공개 프로토타입. 입력 내용은 저장하지 않는다.

파이프라인: 파일(HWP/PDF/DOCX/TXT) 또는 붙여넣기 → 팩트 추출(검문소)
→ 제목 후보 5개 → 본문 생성(성동구 스타일) → 린트 검수(정규식) → 복사/다운로드
"""
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
from datetime import date

import streamlit as st

st.set_page_config(page_title="이지프레스 Easy Press", page_icon="📰", layout="centered")

TODAY = date.today()


# ───────────────────────── 시크릿 ─────────────────────────
def secret(key, default=""):
    # Secrets 파일이 아예 없으면 st.secrets 접근 자체가 예외를 던지므로 폭넓게 방어
    try:
        val = st.secrets[key]
        return val if val is not None else default
    except Exception:
        return default


API_KEY = secret("ANTHROPIC_API_KEY")
PASSCODE = secret("PASSCODE")
MODEL = secret("MODEL", "claude-haiku-4-5-20251001")
GEN_LIMIT = 20  # 세션당 생성 한도


# ───────────────────────── LLM ─────────────────────────
@st.cache_resource
def get_client():
    import anthropic
    return anthropic.Anthropic(api_key=API_KEY)


def ask(prompt, max_tokens=1600):
    msg = get_client().messages.create(
        model=MODEL, max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
    )
    return "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")


def parse_json(raw):
    t = re.sub(r"```json|```", "", raw).strip()
    starts = [i for i in (t.find("{"), t.find("[")) if i >= 0]
    if starts:
        t = t[min(starts):]
    end = max(t.rfind("}"), t.rfind("]"))
    if end > -1:
        t = t[: end + 1]
    return json.loads(t)


# ───────────────────────── 파일 → 텍스트 ─────────────────────────
def docx_to_text(data: bytes) -> str:
    """문단·표 순서를 유지하며 추출. 표는 '셀 | 셀 | 셀' 행으로."""
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(io.BytesIO(data))
    out = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            t = Paragraph(child, doc).text.strip()
            if t:
                out.append(t)
        elif child.tag == qn("w:tbl"):
            tbl = Table(child, doc)
            for row in tbl.rows:
                cells = [" ".join(c.text.split()) for c in row.cells]
                if any(cells):
                    out.append(" | ".join(cells))
            out.append("")
    return "\n".join(out)


def pdf_to_text(data: bytes) -> str:
    """단어 좌표 기반 행·열 재구성(격자선 없는 표 대응) + 격자선 표 전용 추출."""
    import pdfplumber

    out = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages[:15]:
            # 1) 단어 좌표로 본문을 행·열 재구성 — 격자선 없는 표도 열 구분(|) 유지
            rows = {}
            for w in page.extract_words() or []:
                key = round(w["top"] / 3) * 3
                rows.setdefault(key, []).append(w)
            lines = []
            for key in sorted(rows):
                ws = sorted(rows[key], key=lambda x: x["x0"])
                line, prev_x1 = "", None
                for w in ws:
                    if prev_x1 is not None and w["x0"] - prev_x1 > 12:
                        line += " | "
                    elif line:
                        line += " "
                    line += w["text"]
                    prev_x1 = w["x1"]
                if line.strip():
                    lines.append(line.strip())
            if lines:
                out.append("\n".join(lines))
            # 2) 격자선이 있는 표는 구조 그대로 한 번 더(정확도 최상)
            for table in page.extract_tables() or []:
                trows = []
                for row in table:
                    cells = [" ".join((c or "").split()) for c in row]
                    if any(cells):
                        trows.append(" | ".join(cells))
                if trows:
                    out.append("[표 — 구조 인식]\n" + "\n".join(trows))
    return "\n\n".join(out)


def hwp_to_text(data: bytes) -> str:
    """1차 pyhwp(hwp5txt), 실패 시 PrvText 미리보기 스트림 폴백."""
    with tempfile.NamedTemporaryFile(suffix=".hwp", delete=False) as f:
        f.write(data)
        path = f.name
    try:
        text = ""
        if shutil.which("hwp5txt"):
            r = subprocess.run(["hwp5txt", path], capture_output=True, timeout=60)
            if r.returncode == 0:
                text = r.stdout.decode("utf-8", "ignore")
        if len(text.strip()) < 30:
            import olefile
            ole = olefile.OleFileIO(path)
            if ole.exists("PrvText"):
                text = ole.openstream("PrvText").read().decode("utf-16-le", "ignore")
        return text
    finally:
        os.unlink(path)


def hwpx_to_text(data: bytes) -> str:
    import zipfile
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        parts = sorted(n for n in z.namelist() if n.startswith("Contents/section") and n.endswith(".xml"))
        chunks = []
        for n in parts:
            xml = z.read(n).decode("utf-8", "ignore")
            chunks.append(re.sub(r"<[^>]+>", " ", xml))
    return re.sub(r"[ \t]+", " ", "\n".join(chunks)).strip()


def file_to_text(name: str, data: bytes) -> str:
    n = name.lower()
    if n.endswith(".hwp"):
        return hwp_to_text(data)
    if n.endswith(".hwpx"):
        return hwpx_to_text(data)
    if n.endswith(".pdf"):
        return pdf_to_text(data)
    if n.endswith(".docx"):
        return docx_to_text(data)
    return data.decode("utf-8", "ignore")


# ───────────────────────── 프롬프트 ─────────────────────────
def extract_prompt(doc: str) -> str:
    return (
        "너는 성동구청 보도자료 담당자를 돕는 팩트 추출기다. 아래 기획안에서 보도자료 작성에 필요한 정보만 추출해 JSON으로만 응답하라. 코드펜스 금지.\n"
        '스키마: {"유형":"행사 개최"|"사업·정책 발표"|"수상·성과","무엇을":string|null,"언제":string|null,"어디서":string|null,"대상":string|null,"왜":string|null,"어떻게":string|null,"핵심수치":[{"값":string,"근거":string}],"문의처":string|null,"제외한_내부정보":[string],"미확정_신호":[string]}\n'
        "규칙:\n"
        "1) 기획안에 없는 정보는 null. 절대 창작하지 않는다.\n"
        "2) 수치는 원문 표기 그대로. 반올림·환산 금지. 각 수치의 근거 문장을 '근거'에 기록.\n"
        "3) 개인 실명, 휴대전화·내선번호, 예산 산출내역, 결재라인, 계좌번호는 추출하지 말고 '제외한_내부정보'에 항목명만 기록.\n"
        "4) '(안)', '검토 중', '협의 중', '미정' 등 미확정 표현을 발견하면 해당 구절을 '미확정_신호'에 기록.\n"
        "5) 날짜는 '오는 7월 20일부터' 같은 성동구 관행 표기로 변환.\n"
        "6) 문의처는 부서명·대표번호만. 개인 내선뿐이면 null로 두고 제외 목록에 기록.\n"
        "7) 세로줄(|)로 구분된 줄은 표의 행이다. 위아래로 연속된 | 행은 열 위치(몇 번째 칸인지)로 값을 짝지어라. 표 안에서 기간·대상·장소·규모·예산 등 핵심 팩트를 찾아라.\n"
        "8) 요일제·차수처럼 조건에 따라 일정·대상이 나뉘는 표는 대응 관계를 문장으로 풀어 '어떻게'에 기록하라. 예: 출생연도 끝자리 1·6은 4월 27일(월), 2·7은 4월 28일(화)에 신청.\n\n"
        '기획안:\n"""\n' + doc + '\n"""'
    )


def titles_prompt(facts: dict) -> str:
    return (
        '성동구 보도자료 제목 후보 5개를 JSON 배열(["...","..."])로만 응답하라. 코드펜스 금지.\n'
        '제목 공식: "성동구, [핵심 팩트·동사형 야마]⋯ [부연/사업명]" — 쉼표와 말줄임표(⋯)를 쓴다.\n'
        "핵심 수치가 있으면 제목에 우선 배치한다. 과장 수식어(최고·획기적·혁신적 등) 금지, 팩트 기반 강조만. 20~35자 내외.\n"
        "팩트: " + json.dumps(facts, ensure_ascii=False)
    )


def body_prompt(title: str, facts: dict) -> str:
    no_quote = facts.get("인용주체") == "인용 없음"
    subject = facts.get("인용주체") or "유보화 성동구청장"
    lines = [
        "너는 성동구청 보도자료 초안 작성기다. 아래 스타일가이드와 팩트만으로 보도자료 전문을 작성하라. 설명 없이 보도자료만 출력하라.",
        "",
        "[스타일가이드 — 성동구 실제 보도자료 9건에서 추출]",
        '- 리드(첫 문단, 한 문장): "서울 성동구(구청장 유보화)는 [목적·배경 짧게] [무엇을] [언제부터] [한다/추진한다/실시한다/개최한다]고 밝혔다." 보도 날짜는 넣지 않는다.',
    ]
    if no_quote:
        lines.append("- 구조(5~7문단): 리드 → 사업 정의('OO'은 ~하는 사업이다) 또는 배경·필요성 → 세부 내용 2~3문단(문단당 주제 하나) → 계획·기대 → 마지막 문단은 문의·확인 안내('자세한 사항은 OO을 통해 확인할 수 있다.' 꼴, 문의처가 없으면 계획·기대 문단으로 마무리)")
    else:
        lines.append("- 구조(6~7문단): 리드 → 사업 정의('OO'은 ~하는 사업이다) 또는 배경·필요성 → 세부 내용 2~3문단(문단당 주제 하나) → 계획·기대 → 인용(마지막 문단 고정)")
    lines += [
        "- 문단 첫머리 접속: 이에 구는 ~ / 특히 ~ / 또한, ~ / 아울러, ~ / 한편, ~",
        "- 기관 자칭은 첫 등장 이후 '구는'. '우리 구' 금지.",
        "- 문장 내 나열: ▲항목, ▲항목, ▲항목 등",
        "- 숫자는 아라비아, 자릿점 허용. 괄호 수치 보충 가능: 무더위쉼터(204개소)",
        "- 사업명은 '작은따옴표', 외래어 첫 등장은 한글(원어) 병기: 인공지능(AI)",
        "- 날짜에 '오는', '지난' 사용",
        "- 서술어: 운영한다/추진한다/실시한다, ~할 계획이다, ~할 예정이다, ~할 것으로 기대된다",
    ]
    if no_quote:
        lines.append("- 인용 없음 모드: 어떤 인물의 발언이나 인용 문단도 만들지 않는다. 큰따옴표 발언 금지.")
    else:
        lines.append(f'- 인용(마지막 문단): {subject}은 "[사업의 의미·철학]"이라며 "[앞으로의 의지, ~하겠다로 끝맺음]"고 말했다. 이중 인용 구조 고정, 종결은 \'고 말했다\'.')
    lines += [
        "- 과장 수식어 금지. 명령·청유·의문형 금지. '전했다' 금지.",
        "",
        "[철칙]",
        "1) 팩트에 없는 사실을 창작하지 않는다. 비어 있는 필수 정보는 본문에 [OO 확인 필요] 형태로 표기한다.",
        "2) 수치는 팩트의 표기 그대로 사용한다.",
        "3) 발언·인용을 절대 만들지 않는다." if no_quote else "3) 인용문은 팩트의 '인용방향'을 반영해 초안으로 자연스럽게 작성한다.",
        "",
        f"[제목] {title}",
        "[팩트] " + json.dumps(facts, ensure_ascii=False),
        "",
        "출력 형식: 첫 줄에 제목, 빈 줄 하나, 이후 본문 문단들(문단 사이 빈 줄 하나), "
        + ("마지막 문단은 문의 안내 또는 계획·기대." if no_quote else "마지막 문단이 인용."),
    ]
    return "\n".join(lines)


# ───────────────────────── 린트 엔진 ─────────────────────────
def quote_spans(text):
    spans, open_i = [], -1
    for i, ch in enumerate(text):
        if ch in '"\u201c\u201d':
            if open_i < 0:
                open_i = i
            else:
                spans.append((open_i, i))
                open_i = -1
    return spans


def run_lint(text, expect_quote=True):
    spans = quote_spans(text)
    items = []

    def in_spans(i):
        return any(a < i < b for a, b in spans)

    def hits(pattern, exempt=True):
        return [m.group(0) for m in re.finditer(pattern, text) if not (exempt and in_spans(m.start()))]

    def add(rid, name, sev, basis, matches, fixable=False, note=""):
        for m in matches:
            items.append(dict(id=rid, name=name, sev=sev, basis=basis, excerpt=m, fixable=fixable, note=note))

    add("r1", "과장 수식어", "warn", "팩트 기반 강조만('서울 유일' 등)",
        hits(r"최고|최초|획기적|선도적|혁신적|세계적\s*수준|업계\s*선두"), note="팩트로 입증 가능하면 유지")
    add("r2", "간접인용", "error", "보도자료는 발표 주체의 직접 서술", hits(r"라고\s*한다"))
    add("r3", "추측 표현", "warn", "객관 서술 원칙", hits(r"것\s*같다"))
    add("r4", "명령·청유·의문형", "error", "사실의 객관적 전달", hits(r"해라\.|하자\.|까요\?"))
    add("r5", "'전했다' 사용", "error", "성동구 관행: '말했다' 고정",
        hits(r"(?:라고|고)\s*전했다", exempt=False), fixable=True)
    add("r6", "1인칭 표현", "warn", "'구는'으로 자칭", hits(r"우리\s*구|저희"), fixable=True)
    r7 = [m.group(0) for m in re.finditer(r"(\d[\d,]*)\s*~\s*(\d[\d,]*)(만|억|조)", text)
          if not re.search(r"(만|억|조)$", m.group(1))]
    add("r7", "숫자 범위 표기", "warn", "단위어는 앞뒤 숫자 모두", r7, note="'200만~300만원' 형태로")
    add("r8", "접속어 뒤 쉼표", "warn", "그러나·그리고·그러므로·하지만·따라서 뒤 쉼표 금지",
        hits(r"(?:그러나|그리고|그러므로|하지만|따라서),"), fixable=True)
    add("r9", "인용 온점", "error", "'~다\"고 말했다' — 닫는 따옴표 앞 온점 제거",
        hits(r'\.["\u201d]\s*(?:이?라?고|이?라?며|며)', exempt=False), fixable=True)
    long_s = [s.strip()[:28] + "…" for s in re.findall(r"[^.\n]+\.", text) if len(s.strip()) > 100]
    add("r10", "긴 문장", "info", "가독성", long_s, note="두 문장 분리 검토")
    add("r11", "차별·장애 비유", "error", "금기어", hits(r"절름발이|벙어리|장님", exempt=False))
    r12 = []
    for m in re.finditer(r"오는\s*(\d{1,2})월\s*(\d{1,2})일", text):
        if (int(m.group(1)), int(m.group(2))) < (TODAY.month, TODAY.day):
            r12.append(m.group(0))
    add("r12", "시제 불일치", "warn", f"기준일 {TODAY:%Y.%m.%d}", r12, note="'지난'으로 바꾸거나 일정 확인")
    add("r14", "개인 휴대전화 노출", "error", "부서 대표번호로 교체",
        hits(r"01[016789][-\s]?\d{3,4}[-\s]?\d{4}", exempt=False))
    add("r15", "확인 필요 항목", "warn", "배포 전 반드시 채울 것",
        hits(r"\[[^\]]{1,20}확인 필요\]", exempt=False))

    paras = [p.strip() for p in re.split(r"\n{2,}", text) if p.strip()]
    body = paras[1:] if len(paras) > 1 else []
    if body and not body[0].startswith("서울 성동구(구청장 유보화)는"):
        items.append(dict(id="s1", name="리드 공식 이탈", sev="warn",
                          basis="리드는 '서울 성동구(구청장 유보화)는 ~'으로 시작",
                          excerpt=body[0][:24] + "…", fixable=False, note=""))
    has_quote = re.search(r"고\s*말했다", text) is not None
    if expect_quote and not has_quote:
        items.append(dict(id="s2", name="인용 문단 없음", sev="warn",
                          basis="마지막 문단은 이중 인용 + '고 말했다'", excerpt="—", fixable=False, note=""))
    if not expect_quote and has_quote:
        items.append(dict(id="s3", name="인용 없음 모드 위반", sev="warn",
                          basis="발언 문단 감지", excerpt="고 말했다", fixable=False,
                          note="인용 문단을 삭제하거나 인용 주체를 선택하세요"))
    order = {"error": 0, "warn": 1, "info": 2}
    return sorted(items, key=lambda x: order[x["sev"]])


def apply_fixes(text):
    text = re.sub(r"(라고|고)\s*전했다", r"\1 말했다", text)
    text = re.sub(r"우리\s*구", "구", text)
    text = re.sub(r"(그러나|그리고|그러므로|하지만|따라서),\s*", r"\1 ", text)
    text = re.sub(r'\.(["\u201d])(이?라?고|이?라?며|며)', r"\1\2", text)
    return text


# ───────────────────────── 샘플 ─────────────────────────
SAMPLE_PLAN = """청년 AI 실무 아카데미 운영 계획(안)

1. 추진배경
 - 관내 청년의 AI 활용 역량 수요 증가
2. 사업개요
 구분 | 내용
 기간 | 2026. 9. 1. ~ 10. 30. (8주)
 대상 | 관내 거주 청년 60명 (19~39세)
 장소 | 성동구청 대강당 및 온라인 병행
 예산 | 32,000천원 (산출내역: 강사비 18,000천원, 홍보비 8,000천원, 운영비 6,000천원)
3. 세부 추진계획
 - 생성형 AI 업무 활용, 데이터 분석 등 4개 과정
 - 수료생 중 우수 10명 관내 기업 인턴 연계 (협의 중)
4. 행정사항
 - 담당: 일자리정책과 김OO 주무관 (내선 02-2286-XXXX)"""

FIELD_KEYS = ["무엇을", "언제", "어디서", "대상", "왜", "어떻게", "수치", "문의처"]


# ───────────────────────── 상태 ─────────────────────────
ss = st.session_state
# 위젯 키 값은 위젯 생성 전에만 바꿀 수 있으므로 _fill 큐로 다음 런에 주입
for k, v in ss.pop("_fill", {}).items():
    ss[k] = v

ss.setdefault("step", 1)
ss.setdefault("authed", not PASSCODE)
ss.setdefault("gen_count", 0)
ss.setdefault("plan", "")
ss.setdefault("missing", [])
ss.setdefault("excluded", [])
ss.setdefault("pending", [])
ss.setdefault("num_notes", [])
ss.setdefault("titles", [])
ss.setdefault("body", "")
ss.setdefault("last_file", "")
ss.setdefault("f_유형", "사업·정책 발표")
ss.setdefault("f_인용주체", "유보화 성동구청장")
ss.setdefault("f_인용방향", "")
for k in FIELD_KEYS:
    ss.setdefault("f_" + k, "")


def go(step):
    ss["step"] = step
    st.rerun()


def can_generate():
    if ss.gen_count >= GEN_LIMIT:
        st.error(f"세션 생성 한도({GEN_LIMIT}회)에 도달했어요. 페이지를 새로고침하면 초기화됩니다.")
        return False
    return True


def facts_dict():
    f = {k: ss.get("f_" + k, "") for k in FIELD_KEYS}
    f["유형"] = ss.get("f_유형")
    f["인용주체"] = ss.get("f_인용주체")
    if ss.get("f_인용주체") != "인용 없음":
        f["인용방향"] = ss.get("f_인용방향", "")
    return f


# ───────────────────────── 헤더·게이트 ─────────────────────────
c1, c2 = st.columns([1, 5])
if os.path.exists("assets/logo.png"):
    c1.image("assets/logo.png", width=64)
c2.markdown("## 이지프레스 <small style='color:#6E675B'>Easy Press</small>", unsafe_allow_html=True)
c2.caption("성동구 보도자료 초안·검수 도구 (프로토타입)")

if not ss.authed:
    st.text_input("접근코드를 입력하세요", type="password", key="pass_in")
    if st.button("입장", type="primary"):
        if ss.pass_in == PASSCODE:
            ss["authed"] = True
            st.rerun()
        else:
            st.error("접근코드가 올바르지 않아요.")
    st.stop()

if not API_KEY:
    st.error("ANTHROPIC_API_KEY가 설정되지 않았어요. Streamlit Cloud의 Secrets에 추가해 주세요. (README 참고)")
    st.stop()

steps = ["① 팩트 입력", "② 제목 선택", "③ 초안·검수"]
st.markdown(" · ".join(
    f"**{s}**" if i + 1 == ss.step else f"<span style='color:#9a948a'>{s}</span>"
    for i, s in enumerate(steps)), unsafe_allow_html=True)
st.divider()

# ───────────────────────── STEP 1 ─────────────────────────
if ss.step == 1:
    st.info("민원인 실명 등 개인정보와 비공개 문서는 넣지 마세요. 입력 내용은 저장되지 않습니다.")

    up = st.file_uploader("기획안 파일 첨부 (HWP·HWPX·PDF·Word·TXT)", type=["hwp", "hwpx", "pdf", "docx", "txt"])
    if up is not None and ss.get("last_file") != up.name:
        with st.spinner("파일에서 텍스트를 추출하는 중…"):
            try:
                text = file_to_text(up.name, up.read()).strip()
                if len(text) < 30:
                    st.error("텍스트를 거의 추출하지 못했어요. 스캔본일 수 있어요 — 내용을 복사해 붙여넣어 주세요.")
                else:
                    ss["_fill"] = {"plan": text[:12000]}
                    ss["last_file"] = up.name
                    st.rerun()
            except Exception as e:
                st.error(f"파일 처리 실패: {e}")
    if ss.get("last_file"):
        st.caption(f"첨부됨: {ss.last_file} ✓ — 추출된 내용을 확인한 뒤 '팩트 추출'을 눌러주세요.")

    st.text_area("기획안 내용 (붙여넣기 가능)", key="plan", height=220,
                 placeholder="파일을 첨부하거나 기획안 내용을 붙여넣으세요. 예산 산출내역·개인 연락처는 추출 단계에서 자동 제외됩니다.")

    b1, b2 = st.columns(2)
    if b1.button("샘플 기획안 넣어보기"):
        ss["_fill"] = {"plan": SAMPLE_PLAN}
        ss["last_file"] = ""
        st.rerun()
    if b2.button("팩트 추출 → 폼 채우기", type="primary", disabled=not ss.plan.strip()):
        if can_generate():
            with st.spinner("기획안에서 팩트를 추출하는 중…"):
                try:
                    ss.gen_count += 1
                    j = parse_json(ask(extract_prompt(ss.plan), 1300))
                    fill, missing = {}, []
                    for k in ["무엇을", "언제", "어디서", "대상", "왜", "어떻게", "문의처"]:
                        v = j.get(k)
                        if v:
                            fill["f_" + k] = str(v)
                        else:
                            fill["f_" + k] = ""
                            missing.append(k)
                    nums = j.get("핵심수치") or []
                    fill["f_수치"] = ", ".join(str(n.get("값", "")) for n in nums if n.get("값"))
                    if j.get("유형") in ["행사 개최", "사업·정책 발표", "수상·성과"]:
                        fill["f_유형"] = j["유형"]
                    ss["_fill"] = fill
                    ss["missing"] = missing
                    ss["excluded"] = j.get("제외한_내부정보") or []
                    ss["pending"] = j.get("미확정_신호") or []
                    ss["num_notes"] = nums
                    st.rerun()
                except Exception as e:
                    st.error(f"추출 실패: {e}")

    if ss.excluded:
        st.error("기획안에서 제외한 내부 정보 (보도자료에 넣지 않음)\n\n" + "\n".join("· " + x for x in ss.excluded))
    if ss.pending:
        st.warning("확정 전 단계로 보입니다 — 보도 시점을 확인하세요\n\n" + "\n".join("· " + x for x in ss.pending))
    if ss.missing:
        st.warning("확인 필요(기획안에 없어 비워둠): " + ", ".join(ss.missing))

    st.subheader("팩트 확인·수정")
    st.radio("유형", ["행사 개최", "사업·정책 발표", "수상·성과"], key="f_유형", horizontal=True)
    st.text_input("무엇을 (핵심 사업·행사명)", key="f_무엇을", placeholder="예: '반려식물 보급사업' 추진")
    a, b = st.columns(2)
    a.text_input("언제", key="f_언제", placeholder="오는 8월부터 10월까지")
    b.text_input("어디서", key="f_어디서", placeholder="성동구청 대강당")
    st.text_input("대상", key="f_대상", placeholder="관내 거주 19~39세 청년 60명")
    st.text_area("왜 (배경·필요성)", key="f_왜", height=70)
    st.text_area("어떻게 (세부 내용 — 나열은 ▲로 정리됨)", key="f_어떻게", height=70)
    st.text_input("핵심 수치 (쉼표 구분)", key="f_수치", placeholder="60명, 8주, 4개 과정")
    if not ss.f_수치.strip():
        st.caption("⚠ 뉴스가치를 만들 숫자가 없습니다 — 규모·인원·기간 중 하나라도 넣는 걸 권장해요.")
    for n in ss.num_notes:
        if n.get("값"):
            st.caption(f"· {n['값']} ← “{n.get('근거', '')}”")
    st.radio("인용 주체", ["유보화 성동구청장", "구 관계자", "인용 없음"], key="f_인용주체", horizontal=True)
    if ss.f_인용주체 == "인용 없음":
        st.caption("발언 없이 사실만 담는 단신·안내형으로 작성됩니다. 마지막 문단은 문의 안내로 닫히고, 인용 승인 절차 없이 바로 나갈 수 있어요.")
    else:
        st.text_input("인용에 담을 메시지 (선택)", key="f_인용방향", placeholder="예: 청년이 체감하는 실무 교육")
    st.text_input("문의처 (부서·대표번호)", key="f_문의처", placeholder="일자리정책과(☎02-2286-0000)")

    if st.button("제목 후보 5개 생성", type="primary", disabled=not ss.f_무엇을.strip()):
        if can_generate():
            with st.spinner("제목 후보를 뽑는 중…"):
                try:
                    ss.gen_count += 1
                    arr = parse_json(ask(titles_prompt(facts_dict()), 600))
                    ss["titles"] = [str(t) for t in arr][:5]
                    ss["_fill"] = {"title_edit": ss.titles[0] if ss.titles else ""}
                    go(2)
                except Exception as e:
                    st.error(f"제목 생성 실패: {e}")

# ───────────────────────── STEP 2 ─────────────────────────
elif ss.step == 2:
    st.caption("기자는 제목만 보고 5초 안에 버릴지 정합니다. 하나를 고르고, 필요하면 직접 다듬으세요.")

    def _pick():
        ss["_fill"] = {"title_edit": ss.title_radio}

    st.radio("제목 후보", ss.titles, key="title_radio", on_change=_pick)
    st.text_input("선택한 제목 (수정 가능)", key="title_edit")

    a, b = st.columns([1, 2])
    if a.button("← 팩트 수정"):
        go(1)
    if b.button("이 제목으로 본문 생성", type="primary", disabled=not ss.get("title_edit", "").strip()):
        if can_generate():
            with st.spinner("성동구 스타일로 초안을 작성하는 중…"):
                try:
                    ss.gen_count += 1
                    out = ask(body_prompt(ss.title_edit, facts_dict()), 2000).strip()
                    ss["_fill"] = {"body": out}
                    go(3)
                except Exception as e:
                    st.error(f"본문 생성 실패: {e}")

# ───────────────────────── STEP 3 ─────────────────────────
else:
    expect_quote = ss.get("f_인용주체") != "인용 없음"
    if expect_quote:
        st.warning("인용 초안 포함 — 당사자 승인 전 배포 금지")

    st.text_area("초안 (직접 수정 가능)", key="body", height=430)

    lint = run_lint(ss.body, expect_quote)
    fixable_n = sum(1 for i in lint if i["fixable"])
    h1, h2 = st.columns([2, 1])
    h1.subheader(f"검수 리포트 {len(lint)}건")
    if fixable_n and h2.button(f"자동수정 일괄 적용 ({fixable_n})"):
        ss["_fill"] = {"body": apply_fixes(ss.body)}
        st.rerun()
    if not lint:
        st.success("규칙 위반 없음 — 배포 전 팩트와 인용 승인만 확인하세요.")
    icon = {"error": "🟥 오류", "warn": "🟧 경고", "info": "⬜ 참고"}
    for it in lint:
        extra = f" — {it['note']}" if it["note"] else ""
        fx = " · 자동수정 가능" if it["fixable"] else ""
        st.markdown(f"{icon[it['sev']]} **{it['name']}** · “{it['excerpt']}”  \n"
                    f"<small style='color:#6E675B'>{it['basis']}{extra}{fx}</small>", unsafe_allow_html=True)

    st.divider()
    title_line = ss.body.split("\n", 1)[0].strip() if ss.body.strip() else ""
    st.caption("메일 제목 (복사해서 쓰세요)")
    st.code(f"[성동구청 보도자료] {title_line}", language=None)
    st.download_button("초안 텍스트 다운로드", ss.body, file_name="보도자료_초안.txt")

    a, b, c = st.columns(3)
    if a.button("← 제목 다시"):
        go(2)
    if b.button("다시 생성"):
        if can_generate():
            with st.spinner("다시 작성하는 중…"):
                try:
                    ss.gen_count += 1
                    out = ask(body_prompt(ss.get("title_edit", title_line), facts_dict()), 2000).strip()
                    ss["_fill"] = {"body": out}
                    st.rerun()
                except Exception as e:
                    st.error(f"생성 실패: {e}")
    if c.button("새 보도자료"):
        fill = {"plan": "", "body": "", "title_edit": "", "f_유형": "사업·정책 발표",
                "f_인용주체": "유보화 성동구청장", "f_인용방향": ""}
        for k in FIELD_KEYS:
            fill["f_" + k] = ""
        ss["_fill"] = fill
        ss["missing"] = []
        ss["excluded"] = []
        ss["pending"] = []
        ss["num_notes"] = []
        ss["titles"] = []
        ss["last_file"] = ""
        go(1)

# ───────────────────────── 푸터 ─────────────────────────
st.divider()
st.markdown("**문의 및 피드백: 010-8829-5108(정호원)**")
st.caption("본 도구는 성동구 공식 서비스가 아닌 개인 제작 프로토타입입니다. 입력 내용은 저장되지 않으며, "
           "초안의 모든 사실관계는 배포 전 담당자가 확인해야 합니다. "
           "스타일 근거: 성동구 실제 보도자료 9건(2026.6.29~7.9)에서 추출한 하우스 스타일.")
