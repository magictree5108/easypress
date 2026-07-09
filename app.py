# -*- coding: utf-8 -*-
"""
이지프레스 Easy Press — 성동구 보도자료 초안·검수 도구 (Streamlit 버전 v1.0)

기획안(HWP/HWPX/PDF/DOCX/TXT 또는 붙여넣기) → 팩트 추출(검문소) → 제목 후보
→ 본문 생성(성동구 스타일) → 린트 검수 → 복사/다운로드

원칙
1) 입력된 사실만 사용. 없는 정보는 [확인 필요], 창작 금지.
2) 인용문은 초안 — 당사자 승인 전 배포 금지 플래그 상시 표시.
3) 내부 정보(예산 산출내역·개인 실명·내선/휴대전화·결재라인)는 추출에서 제외.
4) 어떤 입력도 저장·로깅하지 않는다.
"""

import io
import re
import json
import zlib
import struct
import datetime
from pathlib import Path

import streamlit as st

# ───────────────────────── 기본 설정 ─────────────────────────

st.set_page_config(page_title="이지프레스 Easy Press", page_icon="📰", layout="centered")


def secret(key, default=""):
    # Secrets 파일이 아예 없으면 st.secrets 접근 자체가 예외를 던지므로 폭넓게 방어
    try:
        val = st.secrets[key]
        return default if val is None else val
    except Exception:
        return default

KST = datetime.timezone(datetime.timedelta(hours=9))
TODAY = datetime.datetime.now(KST).date()
MODEL = secret("MODEL", "claude-haiku-4-5")
MAX_GEN = 20          # 세션당 생성 한도
MAX_PLAN_CHARS = 8000  # 기획안 텍스트 상한
PDF_MAX_PAGES = 15

SAMPLE_PLAN = """청년 AI 실무 아카데미 운영 계획(안)

1. 추진배경
 - 관내 청년의 AI 활용 역량 수요 증가
2. 사업개요
 - 기간: 2026. 9. 1. ~ 10. 30. (8주)
 - 대상: 관내 거주 청년 60명 (19~39세)
 - 장소: 성동구청 대강당 및 온라인 병행
 - 예산: 32,000천원 (산출내역: 강사비 18,000천원, 홍보비 8,000천원, 운영비 6,000천원)
3. 세부 추진계획
 - 생성형 AI 업무 활용, 데이터 분석 등 4개 과정
 - 수료생 중 우수 10명 관내 기업 인턴 연계 (협의 중)
4. 행정사항
 - 담당: 일자리정책과 김OO 주무관 (내선 02-2286-XXXX)"""

FIELDS = ["유형", "무엇을", "언제", "어디서", "대상", "왜", "어떻게", "수치", "인용주체", "인용방향", "문의처"]
FIELD_DEFAULTS = {"유형": "사업·정책 발표", "인용주체": "유보화 성동구청장"}

# ───────────────────────── 세션 상태 ─────────────────────────

ss = st.session_state
_defaults = {
    "step": 1, "plan_text": "", "file_name": "",
    "missing": [], "excluded": [], "pending": [], "num_notes": [],
    "titles": [], "title_final": "", "body": "",
    "gen_count": 0, "authed": False,
}
for k, v in _defaults.items():
    if k not in ss:
        ss[k] = v
for k in FIELDS:
    kk = "f_" + k
    if kk not in ss:
        ss[kk] = FIELD_DEFAULTS.get(k, "")

# 위젯 상태 keep-alive (단계 이동 시 값 유실 방지)
for k in list(ss.keys()):
    if k.startswith("f_") or k in ("plan_text", "title_final"):
        ss[k] = ss[k]


# ───────────────────────── LLM 호출 ─────────────────────────

def ask(prompt: str, max_tokens: int = 1600) -> str:
    import anthropic
    key = secret("ANTHROPIC_API_KEY")
    if not key:
        st.error("ANTHROPIC_API_KEY가 설정되지 않았어요. Streamlit Cloud의 Secrets에 키를 추가해 주세요.")
        st.stop()
    client = anthropic.Anthropic(api_key=key)
    resp = client.messages.create(
        model=MODEL, max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
    )
    return "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")


def parse_json(raw: str):
    t = re.sub(r"```json|```", "", raw).strip()
    starts = [i for i in (t.find("{"), t.find("[")) if i >= 0]
    if starts:
        t = t[min(starts):]
    e = max(t.rfind("}"), t.rfind("]"))
    if e > -1:
        t = t[: e + 1]
    return json.loads(t)


def gen_ok() -> bool:
    if ss.gen_count >= MAX_GEN:
        st.error(f"세션당 생성 한도({MAX_GEN}회)에 도달했어요. 페이지를 새로고침하면 초기화됩니다.")
        return False
    return True


# ───────────────────────── 프롬프트 ─────────────────────────

def extract_prompt(doc: str) -> str:
    return (
        "너는 성동구청 보도자료 담당자를 돕는 팩트 추출기다. 아래 기획안에서 보도자료 작성에 필요한 정보만 추출해 JSON으로만 응답하라. 코드펜스 금지.\n"
        '스키마: {"유형":"행사 개최"|"사업·정책 발표"|"수상·성과","무엇을":string|null,"언제":string|null,"어디서":string|null,"대상":string|null,"왜":string|null,"어떻게":string|null,"핵심수치":[{"값":string,"근거":string}],"문의처":string|null,"기타핵심사실":[string],"제외한_내부정보":[string],"미확정_신호":[string]}\n'
        "규칙:\n"
        "1) 문서에 없는 정보만 null. 문서에 있는 정보를 null로 두는 것은 오답이다 — 스키마에 딱 맞지 않으면 가장 가까운 필드나 '기타핵심사실'에 문장으로 담아라. 절대 창작하지 않는다.\n"
        "2) 수치는 원문 표기 그대로. 반올림·환산 금지. 각 수치의 근거 문장을 '근거'에 기록.\n"
        "3) 개인 실명, 휴대전화·내선번호, 예산 산출내역, 결재라인, 계좌번호는 추출하지 말고 '제외한_내부정보'에 항목명만 기록.\n"
        "4) '(안)', '검토 중', '협의 중', '미정' 등 미확정 표현을 발견하면 해당 구절을 '미확정_신호'에 기록.\n"
        "5) 날짜는 '오는 7월 20일부터' 같은 성동구 관행 표기로 변환.\n"
        "6) 문의처는 부서명·대표번호만. 개인 내선뿐이면 null로 두고 제외 목록에 기록.\n"
        "7) 세로줄(|)로 구분된 줄은 표의 행이다. 위아래로 연속된 | 행은 열 위치(몇 번째 칸인지)로 값을 짝지어라.\n"
        "8) 〈표 풀이〉 블록은 표의 행·열 대응을 코드가 미리 풀어놓은 것이다. 그대로 신뢰하고 활용하라.\n"
        "9) 요일제·차수처럼 조건에 따라 일정·대상이 나뉘는 표는 대응 관계를 문장으로 풀어 '어떻게'에 기록하라. 예: 출생연도 끝자리 1·6은 4월 27일(월), 2·7은 4월 28일(화)에 신청.\n"
        "10) '어떻게'에는 신청 방법, 금액 구분, 지급수단, 사용처와 사용기한 등 실행 세부를 문장으로 압축해 모두 담아라.\n"
        "11) 지원 '대상'(돈·혜택을 받는 사람)과 '사용처'(돈을 쓸 수 있는 곳)를 절대 혼동하지 마라.\n\n"
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
    if no_quote:
        structure = "- 구조(5~7문단): 리드 → 사업 정의('OO'은 ~하는 사업이다) 또는 배경·필요성 → 세부 내용 2~3문단(문단당 주제 하나) → 계획·기대 → 마지막 문단은 문의·확인 안내('자세한 사항은 OO을 통해 확인할 수 있다.' 꼴, 문의처가 없으면 계획·기대 문단으로 마무리)"
        quote_rule = "- 인용 없음 모드: 어떤 인물의 발언이나 인용 문단도 만들지 않는다. 큰따옴표 발언 금지."
        rule3 = "4) 발언·인용을 절대 만들지 않는다."
        out_fmt = "마지막 문단은 문의 안내 또는 계획·기대."
    else:
        structure = "- 구조(6~7문단): 리드 → 사업 정의('OO'은 ~하는 사업이다) 또는 배경·필요성 → 세부 내용 2~3문단(문단당 주제 하나) → 계획·기대 → 인용(마지막 문단 고정)"
        quote_rule = (
            "- 인용(마지막 문단): " + (facts.get("인용주체") or "유보화 성동구청장")
            + '은 "[사업의 의미·철학]"이라며 "[앞으로의 의지, ~하겠다로 끝맺음]"고 말했다. 이중 인용 구조 고정, 종결은 \'고 말했다\'.'
        )
        rule3 = "4) 인용문은 팩트의 '인용방향'을 반영해 자연스럽게 작성한다. 인용방향이 비어 있으면 팩트의 취지에 맞는 초안 인용을 직접 쓴다 — [확인 필요] 금지."
        out_fmt = "마지막 문단이 인용."
    return (
        "너는 성동구청 보도자료 초안 작성기다. 아래 스타일가이드와 팩트만으로 보도자료 전문을 작성하라. 설명 없이 보도자료만 출력하라.\n\n"
        "[스타일가이드 — 성동구 실제 보도자료 9건에서 추출]\n"
        '- 리드(첫 문단, 한 문장): "서울 성동구(구청장 유보화)는 [목적·배경 짧게] [무엇을] [언제부터] [한다/추진한다/실시한다/개최한다]고 밝혔다." 보도 날짜는 넣지 않는다.\n'
        + structure + "\n"
        "- 문단 첫머리 접속: 이에 구는 ~ / 특히 ~ / 또한, ~ / 아울러, ~ / 한편, ~\n"
        "- 기관 자칭은 첫 등장 이후 '구는'. '우리 구' 금지.\n"
        "- 문장 내 나열: ▲항목, ▲항목, ▲항목 등\n"
        "- 숫자는 아라비아, 자릿점 허용. 괄호 수치 보충 가능: 무더위쉼터(204개소)\n"
        "- 사업명은 '작은따옴표', 외래어 첫 등장은 한글(원어) 병기: 인공지능(AI)\n"
        "- 날짜에 '오는', '지난' 사용\n"
        "- 서술어: 운영한다/추진한다/실시한다, ~할 계획이다, ~할 예정이다, ~할 것으로 기대된다\n"
        + quote_rule + "\n"
        "- 과장 수식어 금지. 명령·청유·의문형 금지. '전했다' 금지.\n\n"
        "[철칙]\n"
        "1) 팩트에 없는 사실을 창작하지 않는다. 팩트에 없는 내용은 그 문장 자체를 만들지 마라 — 문단 수를 줄여서라도 있는 팩트만으로 쓴다.\n"
        "2) [OO 확인 필요]는 리드·기간·문의처처럼 문장 성립에 꼭 필요한 자리가 비었을 때만 최소한으로 쓴다. 인용문에는 절대 쓰지 않는다.\n"
        "3) 수치는 팩트의 표기 그대로 사용한다.\n"
        + rule3 + "\n\n"
        "[제목] " + title + "\n"
        "[팩트] " + json.dumps(facts, ensure_ascii=False) + "\n\n"
        "출력 형식: 첫 줄에 제목, 빈 줄 하나, 이후 본문 문단들(문단 사이 빈 줄 하나), " + out_fmt
    )


# ───────────────────────── 파일 → 텍스트 ─────────────────────────

_HWP_CHAR_CONTROLS = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}


def _hwp_records_to_text(buf: bytes) -> str:
    i, n = 0, len(buf)
    parts = []
    while i + 4 <= n:
        (hdr,) = struct.unpack_from("<I", buf, i)
        tag = hdr & 0x3FF
        size = (hdr >> 20) & 0xFFF
        i += 4
        if size == 0xFFF:
            (size,) = struct.unpack_from("<I", buf, i)
            i += 4
        if tag == 67:  # HWPTAG_PARA_TEXT
            j, end = i, min(i + size, n)
            line = []
            while j + 2 <= end:
                (ch,) = struct.unpack_from("<H", buf, j)
                if ch < 32:
                    if ch in (10, 13):
                        line.append("\n")
                    j += 2 if ch in _HWP_CHAR_CONTROLS else 16
                    continue
                line.append(chr(ch))
                j += 2
            t = "".join(line).strip()
            if t:
                parts.append(t)
        i += size
    text = "\n".join(parts)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def hwp_to_text(data: bytes) -> str:
    import olefile
    ole = olefile.OleFileIO(io.BytesIO(data))
    try:
        header = ole.openstream("FileHeader").read()
        (flag,) = struct.unpack_from("<I", header, 36)
        compressed = bool(flag & 1)
        sections = sorted(
            (e for e in ole.listdir() if e[0] == "BodyText" and e[1].startswith("Section")),
            key=lambda e: int(e[1][7:]),
        )
        if not sections:
            raise ValueError("본문을 찾지 못했어요. 암호화되었거나 배포용 문서일 수 있어요.")
        out = []
        for e in sections:
            raw = ole.openstream(e).read()
            if compressed:
                raw = zlib.decompress(raw, -15)
            out.append(_hwp_records_to_text(raw))
        return "\n".join(out)
    finally:
        ole.close()


def hwpx_to_text(data: bytes) -> str:
    import zipfile
    from xml.etree import ElementTree as ET
    zf = zipfile.ZipFile(io.BytesIO(data))
    names = sorted(n for n in zf.namelist() if re.match(r"Contents/section\d+\.xml$", n))
    if not names:
        raise ValueError("HWPX 본문을 찾지 못했어요.")
    out = []
    for n in names:
        root = ET.fromstring(zf.read(n))
        for el in root.iter():
            if el.tag.endswith("}t") and el.text and el.text.strip():
                out.append(el.text.strip())
    return "\n".join(out)


def unfold_table(rows) -> list:
    """표의 행·열 대응을 '헤더: 값' 문장 짝으로 미리 풀어낸다(결정론적).
    첫 행을 헤더로 보고 각 데이터 행을 열 위치로 짝짓는다 — 요일제(2행 표)에 특히 강함."""
    clean = [[re.sub(r"\s+", " ", (c or "")).strip() for c in r] for r in rows]
    clean = [r for r in clean if any(r)]
    if len(clean) < 2:
        return []
    header = clean[0]
    out = []
    for r in clean[1:]:
        pairs = []
        for idx, v in enumerate(r):
            if not v:
                continue
            h = header[idx] if idx < len(header) else ""
            pairs.append(f"{h}: {v}" if h and h != v else v)
        if pairs:
            out.append(" / ".join(pairs))
    return out


def pdf_to_text(data: bytes) -> str:
    import pdfplumber
    out = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages[:PDF_MAX_PAGES]:
            # 1) 단어 좌표 기반 행·열 재구성 — 격자선 없는 표도 열 구분(|) 유지
            rowmap = {}
            for w in page.extract_words() or []:
                key = round(w["top"] / 3) * 3
                rowmap.setdefault(key, []).append(w)
            lines = []
            for key in sorted(rowmap):
                ws = sorted(rowmap[key], key=lambda x: x["x0"])
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
                # 1-1) 연속된 | 행(격자선 없는 표)도 풀이 — 열 개수가 같은 run만
                unfolds, run = [], []

                def _flush(r):
                    if len(r) >= 2 and len({l.count(" | ") for l in r}) == 1:
                        unfolds.extend(unfold_table([l.split(" | ") for l in r]))

                for l in lines + [""]:
                    if l.count(" | ") >= 1:
                        run.append(l)
                    else:
                        _flush(run)
                        run = []
                if unfolds:
                    out.append("〈표 풀이〉\n" + "\n".join(unfolds))
            # 2) 격자선이 있는 표는 구조 그대로 + 풀이까지(정확도 최상)
            for tbl in page.extract_tables() or []:
                rows = []
                for row in tbl:
                    cells = [re.sub(r"\s+", " ", (c or "")).strip() for c in row]
                    if any(cells):
                        rows.append(" | ".join(cells))
                if rows:
                    out.append("〈표〉\n" + "\n".join(rows))
                unf = unfold_table(tbl)
                if unf:
                    out.append("〈표 풀이〉\n" + "\n".join(unf))
    return "\n\n".join(out)


def docx_to_text(data: bytes) -> str:
    import docx
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = docx.Document(io.BytesIO(data))
    out = []
    for child in d.element.body.iterchildren():
        if child.tag == qn("w:p"):
            t = Paragraph(child, d).text.strip()
            if t:
                out.append(t)
        elif child.tag == qn("w:tbl"):
            tbl_rows = []
            for row in Table(child, d).rows:
                cells = [re.sub(r"\s+", " ", c.text).strip() for c in row.cells]
                if any(cells):
                    tbl_rows.append(cells)
            if tbl_rows:
                out.append("〈표〉")
                out.extend(" | ".join(r) for r in tbl_rows)
                unf = unfold_table(tbl_rows)
                if unf:
                    out.append("〈표 풀이〉")
                    out.extend(unf)
            out.append("")
    return "\n".join(out)


def file_to_text(name: str, data: bytes) -> str:
    n = name.lower()
    if n.endswith(".hwp"):
        text = hwp_to_text(data)
    elif n.endswith(".hwpx"):
        text = hwpx_to_text(data)
    elif n.endswith(".pdf"):
        text = pdf_to_text(data)
    elif n.endswith(".docx"):
        text = docx_to_text(data)
    elif n.endswith(".txt"):
        text = data.decode("utf-8", errors="replace")
    elif n.endswith(".doc"):
        raise ValueError("구형 .doc 형식은 지원되지 않아요 — .docx로 저장해 주세요.")
    else:
        raise ValueError("HWP, HWPX, PDF, Word(.docx), TXT 파일만 지원해요.")
    text = re.sub(r"\u0000", "", text or "").strip()
    if len(text) < 30:
        raise ValueError("텍스트를 거의 추출하지 못했어요. 스캔본이거나 이미지 위주 문서일 수 있어요 — 내용을 복사해 붙여넣어 주세요.")
    return text[:MAX_PLAN_CHARS]


# ───────────────────────── 린트 엔진 ─────────────────────────

def _quote_spans(text: str):
    spans, open_i = [], -1
    for i, c in enumerate(text):
        if c in ('"', "\u201c", "\u201d"):
            if open_i < 0:
                open_i = i
            else:
                spans.append((open_i, i))
                open_i = -1
    return spans


def _in_spans(i: int, spans) -> bool:
    return any(a < i < b for a, b in spans)


def run_lint(text: str, expect_quote: bool = True):
    spans = _quote_spans(text)
    items = []

    def add(name, sev, basis, hits, note="", fixable=False, exempt=False):
        for m in hits:
            if exempt and _in_spans(m.start(), spans):
                continue
            items.append({"name": name, "sev": sev, "basis": basis,
                          "excerpt": m.group(0), "note": note, "fixable": fixable})

    add("과장 수식어", "warn", "성동구 관행: 팩트 기반 강조만('서울 유일' 등)",
        re.finditer(r"(최고|최초|획기적|선도적|혁신적|세계적\s*수준|업계\s*선두)", text),
        note="팩트로 입증 가능하면 유지, 아니면 삭제", exempt=True)
    add("간접인용", "error", "보도자료는 발표 주체의 직접 서술",
        re.finditer(r"라고\s*한다", text), exempt=True)
    add("추측 표현", "warn", "객관 서술 원칙",
        re.finditer(r"것\s*같다", text), exempt=True)
    add("명령·청유·의문형", "error", "사실의 객관적 전달 형식",
        re.finditer(r"(해라\.|하자\.|까요\?)", text), exempt=True)
    add("'전했다' 사용", "error", "성동구 관행: '말했다' 고정",
        re.finditer(r"(라고|고)\s*전했다", text), fixable=True)
    add("1인칭 표현", "warn", "성동구 관행: '구는'으로 자칭",
        re.finditer(r"(우리\s*구|저희)", text), fixable=True, exempt=True)

    for m in re.finditer(r"(\d[\d,]*)\s*~\s*(\d[\d,]*)(만|억|조)", text):
        if not re.search(r"(만|억|조)$", m.group(1)):
            items.append({"name": "숫자 범위 표기", "sev": "warn", "basis": "단위어는 앞뒤 숫자에 모두",
                          "excerpt": m.group(0),
                          "note": f"'{m.group(1)}{m.group(3)}~{m.group(2)}{m.group(3)}'처럼",
                          "fixable": False})

    add("접속어 뒤 쉼표", "warn", "그러나·그리고·그러므로·하지만·따라서 뒤 쉼표 금지(또한·아울러는 관행상 허용)",
        re.finditer(r"(그러나|그리고|그러므로|하지만|따라서),", text), fixable=True, exempt=True)
    add("인용 온점", "error", "'~다\"고 말했다' — 닫는 따옴표 앞 온점 제거",
        re.finditer(r'\.["\u201d]\s*(이?라?고|이?라?며|며)', text), fixable=True)

    for s in re.findall(r"[^.\n]+\.", text):
        t = s.strip()
        if len(t) > 100:
            items.append({"name": "긴 문장", "sev": "info", "basis": "가독성(간결함 원칙)",
                          "excerpt": t[:28] + "…", "note": f"{len(t)}자 — 두 문장 분리 검토", "fixable": False})

    add("차별·장애 비유", "error", "금기어",
        re.finditer(r"(절름발이|벙어리|장님)", text))

    for m in re.finditer(r"오는\s*(\d{1,2})월\s*(\d{1,2})일", text):
        if (int(m.group(1)), int(m.group(2))) < (TODAY.month, TODAY.day):
            items.append({"name": "시제 불일치", "sev": "warn", "basis": f"기준일 {TODAY:%Y.%m.%d}",
                          "excerpt": m.group(0), "note": "오늘보다 과거 날짜 — '지난'으로 바꾸거나 일정 확인",
                          "fixable": False})

    add("개인 휴대전화 노출", "error", "개인정보 — 부서 대표번호로 교체",
        re.finditer(r"01[016789][-\s]?\d{3,4}[-\s]?\d{4}", text))
    add("확인 필요 항목", "warn", "창작 금지 원칙 — 배포 전 반드시 채울 것",
        re.finditer(r"\[[^\]]{1,20}확인 필요\]", text))

    paras = [p.strip() for p in re.split(r"\n{2,}", text) if p.strip()]
    body = paras[1:]
    if body and not body[0].startswith("서울 성동구(구청장 유보화)는"):
        items.append({"name": "리드 공식 이탈", "sev": "warn",
                      "basis": "리드는 '서울 성동구(구청장 유보화)는 ~'으로 시작",
                      "excerpt": body[0][:24] + "…", "note": "", "fixable": False})
    has_quote = bool(re.search(r"고\s*말했다", text))
    if expect_quote and not has_quote:
        items.append({"name": "인용 문단 없음", "sev": "warn",
                      "basis": "마지막 문단은 이중 인용 + '고 말했다'", "excerpt": "—", "note": "", "fixable": False})
    if not expect_quote and has_quote:
        items.append({"name": "인용 없음 모드 위반", "sev": "warn",
                      "basis": "인용 없음으로 설정했는데 발언 문단이 감지됨", "excerpt": "고 말했다",
                      "note": "발언 문단을 삭제하거나 인용 주체를 선택하세요", "fixable": False})

    order = {"error": 0, "warn": 1, "info": 2}
    return sorted(items, key=lambda x: order[x["sev"]])


def apply_fixes(t: str) -> str:
    t = re.sub(r"(라고|고)\s*전했다", r"\1 말했다", t)
    t = re.sub(r"우리\s*구", "구", t)
    t = re.sub(r"(그러나|그리고|그러므로|하지만|따라서),\s*", r"\1 ", t)
    t = re.sub(r'\.(["\u201d])(고|라고|라며|이라며|며)', r"\1\2", t)
    return t

# ───────────────────────── 헤더 ─────────────────────────

_logo = Path("assets/logo.png")
c1, c2 = st.columns([1, 6])
with c1:
    if _logo.exists():
        st.image(str(_logo), width=56)
with c2:
    st.markdown("### 이지프레스 &nbsp;<span style='font-size:0.6em;color:#6E675B'>Easy Press</span>", unsafe_allow_html=True)
    st.caption("성동구 보도자료 초안·검수 · 개인 제작 프로토타입")

_steps = ["① 팩트 입력", "② 제목 선택", "③ 초안·검수"]
st.markdown(" · ".join(f"**{s}**" if i + 1 == ss.step else f"<span style='color:#999'>{s}</span>" for i, s in enumerate(_steps)), unsafe_allow_html=True)
st.divider()


def facts_dict() -> dict:
    return {k: ss["f_" + k] for k in FIELDS}


# ───────────────────────── 1단계: 팩트 입력 ─────────────────────────

if ss.step == 1:
    st.markdown("**기획안 넣기 — 파일 업로드 또는 붙여넣기**")
    up = st.file_uploader("기획안 파일 (HWP·HWPX·PDF·DOCX·TXT)", type=["hwp", "hwpx", "pdf", "docx", "txt"])
    b1, b2 = st.columns(2)
    if b1.button("파일에서 텍스트 추출", disabled=up is None, use_container_width=True):
        try:
            ss.plan_text = file_to_text(up.name, up.getvalue())
            ss.file_name = up.name
            st.toast(f"{up.name} 추출 완료")
        except Exception as e:
            st.error(str(e))
    if b2.button("샘플 기획안 넣어보기", use_container_width=True):
        ss.plan_text = SAMPLE_PLAN
        ss.file_name = ""
    if ss.file_name:
        st.caption(f"첨부됨: {ss.file_name} ✓")

    st.text_area("기획안 텍스트 (확인·수정 가능)", key="plan_text", height=220,
                 placeholder="파일을 업로드하거나 한글(HWP) 기획안 내용을 복사해 붙여넣으세요. 예산 산출내역·개인 연락처는 자동으로 제외됩니다.")
    st.caption("업로드는 폼을 대체하지 않고 폼을 채웁니다. 추출된 텍스트를 확인한 뒤 아래 버튼을 눌러주세요"
               f"(긴 문서는 앞 {MAX_PLAN_CHARS:,}자 사용). 민원인 실명 등 개인정보·비공개 문서는 넣지 마세요.")

    if st.button("팩트 추출 → 폼 채우기", type="primary", disabled=not ss.plan_text.strip(), use_container_width=True):
        if gen_ok():
            try:
                with st.spinner("기획안에서 팩트를 추출하는 중…"):
                    j = parse_json(ask(extract_prompt(ss.plan_text)))
                ss.gen_count += 1
                miss = []

                def pick(v, k):
                    if v in (None, ""):
                        miss.append(k)
                        return ""
                    return str(v)

                if j.get("유형") in ("행사 개최", "사업·정책 발표", "수상·성과"):
                    ss.f_유형 = j["유형"]
                for k in ("무엇을", "언제", "어디서", "대상", "왜", "어떻게"):
                    ss["f_" + k] = pick(j.get(k), k)
                nums = j.get("핵심수치") or []
                ss.f_수치 = ", ".join(str(n.get("값", "")) for n in nums if n.get("값"))
                ss.num_notes = nums
                ss.f_문의처 = pick(j.get("문의처"), "문의처")
                extra = j.get("기타핵심사실") or []
                if extra:
                    joined = " / ".join(str(x) for x in extra if x)
                    if joined:
                        ss.f_어떻게 = (ss.f_어떻게 + " / " + joined) if ss.f_어떻게 else joined
                        if "어떻게" in miss:
                            miss.remove("어떻게")
                ss.missing = miss
                ss.excluded = j.get("제외한_내부정보") or []
                ss.pending = j.get("미확정_신호") or []
                st.rerun()
            except Exception as e:
                st.error(f"추출 실패: {e}")

    if ss.excluded:
        st.error("기획안에서 제외한 내부 정보 (보도자료에 넣지 않음)\n\n- " + "\n- ".join(map(str, ss.excluded)))
    if ss.pending:
        st.warning("확정 전 단계로 보입니다 — 보도 시점을 확인하세요\n\n- " + "\n- ".join(map(str, ss.pending)))

    st.divider()
    st.radio("유형", ["행사 개최", "사업·정책 발표", "수상·성과"], key="f_유형", horizontal=True)

    def field(k, ph, area=False):
        label = ("🔴 " if k in ss.missing and not ss["f_" + k] else "") + k + (" · 확인 필요" if k in ss.missing and not ss["f_" + k] else "")
        if area:
            st.text_area(label, key="f_" + k, placeholder=ph, height=80)
        else:
            st.text_input(label, key="f_" + k, placeholder=ph)

    field("무엇을", "핵심 사업·행사명 (예: '반려식물 보급사업' 추진)")
    field("언제", "오는 8월부터 10월까지 / 오는 7월 20일 오후 2시")
    field("어디서", "성동구청 대강당, 온라인 병행")
    field("대상", "노인맞춤돌봄서비스 이용 65세 이상 어르신 118명")
    field("왜", "추진 배경·필요성", area=True)
    field("어떻게", "세부 내용·운영 방식 (나열은 ▲로 정리됨)", area=True)

    st.text_input("핵심 수치", key="f_수치", placeholder="60명, 8주, 4개 과정 (쉼표로 구분)")
    if not ss.f_수치.strip():
        st.warning("뉴스가치를 만들 숫자가 없습니다 — 규모·인원·기간 중 하나라도 넣는 걸 권장해요.")
    for n in ss.num_notes:
        if n.get("값") and n.get("근거"):
            st.caption(f"· {n['값']} ← \"{n['근거']}\"")

    st.radio("인용 주체", ["유보화 성동구청장", "구 관계자", "인용 없음"], key="f_인용주체", horizontal=True)
    if ss.f_인용주체 == "인용 없음":
        st.caption("발언 없이 사실만 담는 단신·안내형으로 작성됩니다. 마지막 문단은 문의 안내로 닫히고, 인용 승인 절차 없이 바로 나갈 수 있어요.")
    else:
        st.text_input("인용에 담을 메시지 (선택)", key="f_인용방향", placeholder="예: 청년이 체감하는 실무 교육, 현장 목소리 반영")

    field("문의처", "일자리정책과(☎02-2286-0000)")

    st.caption("입력한 사실만 사용합니다. 비어 있는 정보는 초안에 [확인 필요]로 표시되고, 창작하지 않습니다.")
    if st.button("제목 후보 5개 생성", type="primary", disabled=not ss.f_무엇을.strip(), use_container_width=True):
        if gen_ok():
            try:
                with st.spinner("제목 후보를 뽑는 중…"):
                    arr = parse_json(ask(titles_prompt(facts_dict()), max_tokens=800))
                ss.gen_count += 1
                ss.titles = [str(t) for t in arr][:5]
                ss.title_final = ss.titles[0] if ss.titles else ""
                ss.step = 2
                st.rerun()
            except Exception as e:
                st.error(f"제목 생성 실패: {e}")

# ───────────────────────── 2단계: 제목 선택 ─────────────────────────

elif ss.step == 2:
    st.caption("기자는 제목만 보고 5초 안에 버릴지 정합니다. 하나를 고르고, 필요하면 직접 다듬으세요.")

    def _pick_title():
        ss.title_final = ss.title_pick

    st.radio("제목 후보", ss.titles, key="title_pick", on_change=_pick_title)
    st.text_input("선택한 제목 (수정 가능)", key="title_final")

    c1, c2 = st.columns([1, 2])
    if c1.button("← 팩트 수정", use_container_width=True):
        ss.step = 1
        st.rerun()
    if c2.button("이 제목으로 본문 생성", type="primary", disabled=not ss.title_final.strip(), use_container_width=True):
        if gen_ok():
            try:
                with st.spinner("성동구 스타일로 초안을 작성하는 중…"):
                    ss.body = ask(body_prompt(ss.title_final, facts_dict()), max_tokens=1600).strip()
                ss.gen_count += 1
                ss.step = 3
                st.rerun()
            except Exception as e:
                st.error(f"본문 생성 실패: {e}")

# ───────────────────────── 3단계: 초안·검수 ─────────────────────────

else:
    paras = [p.strip() for p in re.split(r"\n{2,}", ss.body) if p.strip()]
    doc_title = paras[0] if paras else ss.title_final
    body_paras = paras[1:]
    expect_quote = ss.f_인용주체 != "인용 없음"
    has_quote = bool(re.search(r"고\s*말했다", ss.body))

    if has_quote:
        st.warning("인용 초안 포함 — 당사자 승인 전 배포 금지")

    st.markdown("#### " + doc_title)
    for p in body_paras:
        st.markdown(p)

    st.caption("아래 코드 상자 우측 상단 아이콘으로 전문을 복사할 수 있어요.")
    st.code(ss.body, language=None)
    st.code(f"[성동구청 보도자료] {doc_title}", language=None)
    st.download_button("TXT로 다운로드", data=ss.body, file_name=f"보도자료_{TODAY:%m%d}.txt", use_container_width=True)

    st.divider()
    lint = run_lint(ss.body, expect_quote)
    fixable = sum(1 for i in lint if i["fixable"])
    lc1, lc2 = st.columns([2, 1])
    lc1.markdown(f"**검수 리포트 {len(lint)}건**")
    if fixable and lc2.button(f"자동수정 일괄 적용 ({fixable})", use_container_width=True):
        ss.body = apply_fixes(ss.body)
        st.rerun()

    if not lint:
        st.success("규칙 위반 없음 — 배포 전 팩트와 인용 승인만 확인하세요.")
    else:
        icon = {"error": "🔴", "warn": "🟡", "info": "⚪"}
        for it in lint:
            extra = (" — " + it["note"]) if it["note"] else ""
            fx = " · 자동수정 가능" if it["fixable"] else ""
            st.markdown(f"{icon[it['sev']]} **{it['name']}** · `{it['excerpt']}`  \n"
                        f"<span style='font-size:0.85em;color:#6E675B'>{it['basis']}{extra}{fx}</span>",
                        unsafe_allow_html=True)

    st.divider()
    n1, n2, n3 = st.columns(3)
    if n1.button("← 제목 다시", use_container_width=True):
        ss.step = 2
        st.rerun()
    if n2.button("다시 생성", use_container_width=True):
        if gen_ok():
            try:
                with st.spinner("초안을 다시 작성하는 중…"):
                    ss.body = ask(body_prompt(ss.title_final, facts_dict()), max_tokens=1600).strip()
                ss.gen_count += 1
                st.rerun()
            except Exception as e:
                st.error(f"본문 생성 실패: {e}")
    if n3.button("새 보도자료", use_container_width=True):
        for k in list(ss.keys()):
            if k.startswith("f_") or k in ("plan_text", "file_name", "missing", "excluded", "pending",
                                           "num_notes", "titles", "title_final", "body", "title_pick"):
                del ss[k]
        ss.step = 1
        st.rerun()

# ───────────────────────── 푸터 ─────────────────────────

st.divider()
st.markdown("**문의 및 피드백: 010-8829-5108(정호원)**")
st.caption("본 도구는 성동구 공식 서비스가 아닌 개인 제작 프로토타입입니다. "
           "민원인 실명 등 개인정보·비공개 문서는 입력하지 마세요. 입력 내용은 저장되지 않습니다. "
           "생성된 초안의 모든 사실관계는 배포 전 담당자가 확인해야 합니다.")
