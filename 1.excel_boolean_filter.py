# -*- coding: utf-8 -*-
"""
엑셀 불리언 검색 FILTER 함수식 생성기 (Streamlit 웹앱)

실행 방법
    pip install streamlit
    streamlit run excel_boolean_filter.py

검색어 규칙 (한 칸 안에서)
    *   : AND  (예: 블록체인*보안)
    +   : OR   (예: AI+인공지능)
    ~   : NOT  (예: ~빅데이터)
    ( ) : 묶음 (예: (AI+인공지능)*~빅데이터)
    우선순위는 ~ > * > + 이고, 칸과 칸은 AND로 연결된다.

출력되는 함수식
    셀 참조형 : 엑셀 셀(검색칸)에 적은 검색어를 엑셀이 직접 해석한다. 검색어를 바꿔도 함수식은 그대로 둔다.
    값 고정형 : 이 화면에 적은 검색어를 함수식 안에 그대로 넣는다.

경로 입력
    엑셀 범위 (예: A2:A100, Sheet1!A2:A100)
    쉼표로 여러 범위를 적으면 이어 붙여서 검색 (예: A2:A100, C2:C100)
"""
import streamlit as st

OPERATORS = "+*~()"

# 셀 참조형에서 검색칸의 문자열을 해석하는 재귀 LAMBDA.
# 괄호 밖의 + → 괄호 밖의 * → 맨 앞의 ~ → 바깥 괄호 순으로 쪼개고, 남은 단어를 SEARCH 한다.
# (LET 이름으로 r, c 는 쓸 수 없어 세 글자 이름을 사용한다.)
PARSER_LINES = [
    "LAMBDA(self,qry,rng,LET(",
    "txt,TRIM(qry),",
    "cnt,LEN(txt),",
    "chr,MID(txt,SEQUENCE(MAX(cnt,1)),1),",
    'lvl,SCAN(0,chr,LAMBDA(acc,x,acc+(x="(")-(x=")"))),',
    'pOr,IFERROR(MATCH(1,(chr="+")*(lvl=0),0),0),',
    'pAnd,IFERROR(MATCH(1,(chr="*")*(lvl=0),0),0),',
    "IF(pOr,SIGN(self(self,LEFT(txt,pOr-1),rng)+self(self,MID(txt,pOr+1,cnt),rng)),",
    "IF(pAnd,self(self,LEFT(txt,pAnd-1),rng)*self(self,MID(txt,pAnd+1,cnt),rng),",
    'IF(LEFT(txt)="~",1-self(self,MID(txt,2,cnt),rng),',
    'IF((LEFT(txt)="(")*(RIGHT(txt)=")"),self(self,MID(txt,2,cnt-2),rng),',
    "--ISNUMBER(SEARCH(txt,rng))))))))",
]


# ─────────────────────────── 검색어 해석 ───────────────────────────
class QueryError(ValueError):
    """검색어 문법 오류."""


def tokenize(text: str) -> list[str]:
    """검색어를 연산자와 단어로 나눈다."""
    tokens, word = [], ""
    for ch in text:
        if ch in OPERATORS:
            if word.strip():
                tokens.append(word.strip())
            tokens.append(ch)
            word = ""
        else:
            word += ch
    if word.strip():
        tokens.append(word.strip())
    return tokens


def join(kind: str, items: list) -> tuple:
    """같은 연산이 겹치면 한 단계로 펼친다. 예: (A+B)+C → A+B+C"""
    if len(items) == 1:
        return items[0]
    flat = []
    for item in items:
        flat += item[1] if item[0] == kind else [item]
    return (kind, flat)


def parse(text: str) -> tuple:
    """검색어를 트리로 변환한다. 노드: ("term", 단어) / ("not", 노드) / ("and"|"or", [노드...])"""
    tokens = tokenize(text)
    pos = 0

    def peek():
        return tokens[pos] if pos < len(tokens) else None

    def take() -> str:
        nonlocal pos
        pos += 1
        return tokens[pos - 1]

    def chain(operator: str, kind: str, operand) -> tuple:
        items = [operand()]
        while peek() == operator:
            take()
            items.append(operand())
        return join(kind, items)

    def expr() -> tuple:
        return chain("+", "or", term)

    def term() -> tuple:
        return chain("*", "and", factor)

    def factor() -> tuple:
        token = peek()
        if token is None:
            raise QueryError("연산자 뒤에 검색어가 없습니다.")
        if token in ")+*":
            raise QueryError(f"'{token}' 앞에 검색어가 없습니다.")
        take()
        if token == "~":
            return ("not", factor())
        if token == "(":
            node = expr()
            if peek() != ")":
                raise QueryError("닫는 괄호 ')'가 없습니다.")
            take()
            return node
        return ("term", token)

    node = expr()
    if peek() == ")":
        raise QueryError("여는 괄호 '('가 없습니다.")
    if peek() is not None:
        raise QueryError(f"'{peek()}' 앞에 연산자(*, +)가 없습니다.")
    return node


# ─────────────────────────── 함수식 생성 로직 ───────────────────────────
def split_csv(text: str) -> list[str]:
    """쉼표로 구분된 문자열을 리스트로 변환한다."""
    return [t.strip() for t in (text or "").replace("，", ",").split(",") if t.strip()]


def quote(text: str) -> str:
    """엑셀 문자열 리터럴로 변환한다 (큰따옴표 이스케이프)."""
    return '"' + text.replace('"', '""') + '"'


def range_expr(path: str) -> str:
    """경로 입력을 엑셀 범위식으로 변환한다. 여러 범위는 공백을 두고 이어 붙인다."""
    return '&" "&'.join(split_csv(path))


def node_cond(node: tuple, rng: str) -> str:
    """트리를 엑셀 조건식으로 변환한다."""
    kind, body = node
    if kind == "term":
        return f"ISNUMBER(SEARCH({quote(body)},{rng}))"
    if kind == "not":
        return f"NOT({node_cond(body, rng)})"
    parts = [node_cond(n, rng) for n in body]
    if kind == "and":
        return "(" + "*".join(parts) + ")"
    return "((" + "+".join(parts) + ")>0)"


def node_text(node: tuple, top: bool = True) -> str:
    """트리를 AND / OR / NOT 표기로 변환한다."""
    kind, body = node
    if kind == "term":
        return body
    if kind == "not":
        return "NOT " + node_text(body, top=False)
    text = f" {kind.upper()} ".join(node_text(n, top=False) for n in body)
    return text if top else f"({text})"


def filter_formula(ret: str, conds: list[str], empty_msg: str, multiline: bool,
                   indent: str = "") -> str:
    """조건식들을 AND로 묶어 FILTER 식을 만든다."""
    if not multiline:
        return f"FILTER({ret}," + "*".join(conds) + f",{quote(empty_msg)})"
    pad = indent + "  "
    return (f"FILTER({ret},\n{pad}" + f"\n{pad}*".join(conds)
            + f",\n{pad}{quote(empty_msg)})")


def check_common(v: dict, active: list[dict], empty_error: str) -> list[str]:
    errors = []
    if not v["ret_path"].strip():
        errors.append("결과 반환 범위가 비어 있습니다.")
    if not active:
        errors.append(empty_error)
    for f in active:
        if not range_expr(f["path"]):
            errors.append(f"{f['label']}의 검색 대상 범위가 비어 있습니다.")
    return errors


def build_preview(v: dict) -> tuple[str, list[str]]:
    """입력된 검색어를 (Query Preview, 오류목록)으로 변환한다."""
    lines, errors = [], []
    for f in v["fields"]:
        if not f["query"].strip():
            continue
        try:
            lines.append(f"({f['label']}: {node_text(parse(f['query']))})")
        except QueryError as e:
            errors.append(f"{f['label']}: {e}")
    return "\nAND ".join(lines), errors


def build_dynamic(v: dict) -> tuple[str, list[str]]:
    """셀 참조형: 검색칸 셀의 검색어를 엑셀이 해석하는 함수식."""
    active = [f for f in v["fields"] if f["cell"].strip()]
    errors = check_common(v, active, "검색칸 셀을 하나 이상 입력하세요.")
    if errors:
        return "", errors

    conds = [f"BSEARCH(BSEARCH,{f['cell'].strip()},{range_expr(f['path'])})" for f in active]
    ret = v["ret_path"].strip()
    if v["multiline"]:
        parser = "\n    ".join(PARSER_LINES)
        body = filter_formula(ret, conds, v["empty_msg"], True, indent="  ")
        return f"=LET(\n  BSEARCH,{parser},\n  {body})", []
    body = filter_formula(ret, conds, v["empty_msg"], False)
    return "=LET(BSEARCH," + "".join(PARSER_LINES) + f",{body})", []


def build_static(v: dict) -> tuple[str, list[str]]:
    """값 고정형: 입력한 검색어를 그대로 넣은 함수식."""
    active = [f for f in v["fields"] if f["query"].strip()]
    errors = check_common(v, active, "검색어를 하나 이상 입력하세요.")
    conds = []
    for f in active:
        try:
            conds.append(node_cond(parse(f["query"]), range_expr(f["path"])))
        except QueryError as e:
            errors.append(f"{f['label']}: {e}")
    if errors:
        return "", errors
    return "=" + filter_formula(v["ret_path"].strip(), conds, v["empty_msg"], v["multiline"]), []


# ─────────────────────────────── 화면 ───────────────────────────────
DEFAULT_FIELDS = [
    # (항목명, 검색어, 검색 대상 범위, 검색칸 셀)
    ("제목", "(AI+인공지능)*~빅데이터", "A2:A100", "E1"),
    ("저자", "~홍길동", "B2:B100", "E2"),
    ("초록", "", "C2:C100", "E3"),
]


def field(i: int) -> None:
    """[항목명 | 검색어 | 검색 대상 범위 | 검색칸 셀] 한 줄을 그린다."""
    if i < len(DEFAULT_FIELDS):
        label, query, path, cell = DEFAULT_FIELDS[i]
    else:
        label, query, path, cell = f"검색칸{i + 1}", "", "", ""
    c1, c2, c3, c4 = st.columns([1, 3, 1.5, 1])
    c1.text_input("항목명", value=label, key=f"label{i}")
    c2.text_input("검색어", value=query, key=f"query{i}", placeholder="예: (AI+인공지능)*~빅데이터")
    c3.text_input("검색 대상 범위", value=path, key=f"path{i}", placeholder="예: A2:A100")
    c4.text_input("검색칸 셀", value=cell, key=f"cell{i}", placeholder="예: E1")


def show(formula: str, errors: list[str]) -> None:
    if errors:
        for e in errors:
            st.error(e)
        return
    st.code(formula, language="excel")
    if len(formula) > 8192:
        st.warning(f"함수식 길이({len(formula):,}자)가 엑셀 한도(8,192자)를 넘습니다.")


def main() -> None:
    st.set_page_config(page_title="Query Builder for EPTS", layout="wide")
    st.title("Query Builder for EPTS")
    st.caption("한 칸 안에서는 * (AND), + (OR), ~ (NOT), 괄호를 사용합니다. 우선순위는 ~ > * > + 이고, "
               "칸과 칸은 AND로 연결됩니다. 모든 범위는 행 수가 같아야 합니다.")

    st.subheader("검색어 입력")
    count = st.number_input("검색 칸 수", min_value=1, max_value=10, value=2, key="count")
    for i in range(count):
        field(i)

    st.subheader("출력 설정")
    c1, c2 = st.columns(2)
    c1.text_input("결과 반환 범위", value="A2:C100", key="ret_path")
    c2.text_input("결과 없음 문구", value="검색된 논문이 없습니다.", key="empty_msg")
    st.checkbox("함수식 줄바꿈 표시", value=True, key="multiline")

    values = {k: st.session_state[k] for k in ("ret_path", "empty_msg", "multiline")}
    values["fields"] = [
        {k: st.session_state[f"{k}{i}"] for k in ("label", "query", "path", "cell")}
        for i in range(count)
    ]
    for i, f in enumerate(values["fields"]):
        f["label"] = f["label"].strip() or f"검색칸{i + 1}"

    preview, preview_errors = build_preview(values)
    st.subheader("연산식 Query Preview")
    st.code(preview or " ", language=None)
    for e in preview_errors:
        st.error(e)

    st.subheader("함수식 Search Filter")
    tab1, tab2 = st.tabs(["셀 참조형 (엑셀 검색칸 연동)", "값 고정형 (검색어 내장)"])
    with tab1:
        st.caption("검색칸 셀에 검색어를 적으면 결과가 바로 바뀝니다. 빈 검색칸은 조건에서 빠집니다. "
                   "Microsoft 365 또는 Excel 2024 (LAMBDA 지원 버전)에서 동작합니다.")
        show(*build_dynamic(values))
    with tab2:
        st.caption("위에 입력한 검색어가 함수식에 그대로 들어갑니다. "
                   "FILTER를 지원하는 Excel 2021 이상에서 동작합니다.")
        show(*build_static(values))


if __name__ == "__main__":
    main()
