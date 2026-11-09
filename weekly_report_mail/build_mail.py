#!/usr/bin/env python3
"""플랫폼기획부 주간보고 메일 생성기.

주간 보고 회의록 최종본(구글 시트 탭 / xlsx / 지난주 메일 .eml)을 읽어
메일 작성 창에 그대로 붙여넣을 수 있는 HTML 본문과 제목·텍스트본을 만든다.

사용법
  # 1) 회의록(xlsx로 내려받은 '플랫폼기획부 주간 보고')에서 바로 생성
  python3 build_mail.py from-xlsx 주간보고.xlsx --sheet "주간 보고(11/9)" -o out/261109

  # 2) 지난주 메일(.eml)을 JSON으로 풀어 이번 주 수정의 출발점으로 사용
  python3 build_mail.py from-eml 지난주.eml -o out/base.json

  # 3) JSON(직접 수정했거나 Claude가 회의록에서 채운 것)으로 메일 생성
  python3 build_mail.py build report.json -o out/261109

결과물 (출력 폴더)
  mail.html    브라우저로 열고 [본문 복사] → 메일 작성 창에 Ctrl+V (서식 유지)
  body.html    본문 HTML 조각 (HTML 편집 모드 붙여넣기용)
  mail.txt     텍스트 전용 본문
  subject.txt  메일 제목
  report.json  생성에 사용한 데이터 (다음 주 출발점으로 재사용 가능)

외부 라이브러리: from-xlsx 에만 openpyxl 필요. 나머지는 표준 라이브러리만 사용.
"""
from __future__ import annotations

import argparse
import datetime as dt
import email
import html
import json
import re
import sys
from email import policy
from html.parser import HTMLParser
from pathlib import Path

HERE = Path(__file__).resolve().parent
JIRA_BASE = "https://kidari.atlassian.net/browse/"
FONT = "'맑은 고딕', 'Malgun Gothic', 돋움, dotum, AppleGothic, arial, Helvetica, sans-serif"

# ---------------------------------------------------------------------------
# 표 구성 (기존 주간보고 메일과 동일한 11개 열)
# ---------------------------------------------------------------------------
FIELDS = ["no", "type", "jira", "platform", "project", "detail",
          "schedule", "status", "date", "owner", "note"]
# DATA 표처럼 한 프로젝트에 여러 단계가 있는 경우, 아래 필드는 단계(steps)별로 행이 나뉜다.
STEP_FIELDS = ["detail", "schedule", "status", "date", "owner", "note"]
COL_WIDTHS = [57, 74, 120, 70, 210, 410, 202, 103, 102, 86, 353]
CENTER = {"no", "type", "jira", "platform", "status", "date", "owner"}

DEFAULT_HEADERS = {
    "RELEASE": ["NO", "구분", "JIRA KEY", "대상 플랫폼", "Project", "상세 내용",
                "일정(KST)", "진행 상태", "배포일", "담당자", "특이사항"],
    "PROJECT": ["NO", "구분", "JIRA KEY", "대상 플랫폼", "Project", "상세 내용",
                "일정", "진행 상태", "배포 예정일", "담당자", "특이사항"],
    "TASK": ["NO", "구분", "JIRA KEY", "대상 플랫폼", "Project", "상세 내용",
             "일정", "진행 상태", "배포 예정일", "담당자", "특이 사항\n(외부 공유용)"],
    "UX/UI DESIGN": ["NO", "구분", "JIRA KEY", "대상 플랫폼", "Project", "상세 내용",
                     "일정", "진행 상태", "배포 예정일", "담당자", "특이 사항\n(외부 공유용)"],
    "DATA": ["NO", "구분", "JIRA KEY", "대상 플랫폼", "Project", "상세 내용",
             "일정(KST)", "진행 상태", "배포일", "담당자", "특이 사항\n(외부 공유용)"],
}
SECTION_ORDER = ["RELEASE", "PROJECT", "TASK", "UX/UI DESIGN", "DATA"]

# 헤더 문구 → 필드 (시트/메일에서 열 위치가 바뀌어도 헤더 이름으로 찾는다)
HEADER_ALIASES = {
    "no": "no", "구분": "type", "jirakey": "jira", "jira": "jira",
    "대상플랫폼": "platform", "플랫폼": "platform", "project": "project", "프로젝트": "project",
    "상세내용": "detail", "일정": "schedule", "일정(kst)": "schedule",
    "진행상태": "status", "상태": "status", "배포일": "date", "배포예정일": "date",
    "담당자": "owner", "특이사항": "note", "특이사항(외부공유용)": "note", "비고": "note",
}

# 진행 상태별 글자색 (기존 메일 서식 기준. 너무 옅은 색은 메일에서 읽히도록 진하게 조정)
STATUS_COLORS = {
    "배포": "#ff0000", "배포대기": "#ff0000", "배포 대기": "#ff0000",
    "개발": "#b10202", "기획": "#473821", "디자인": "#0a53a8", "QA": "#2b6777",
    "완료": "#215a6c", "진행": "#5a3286", "홀딩": "#666666", "대기": "#000000",
}
HEADER_BG = "#000000"
NOTE_HEADER_BG = "#990000"


def norm_header(text: str) -> str:
    return re.sub(r"\s+", "", (text or "")).lower()


def header_to_field(text: str) -> str | None:
    key = norm_header(text)
    if key in HEADER_ALIASES:
        return HEADER_ALIASES[key]
    for alias, field in HEADER_ALIASES.items():
        if key.startswith(alias) and len(alias) > 2:
            return field
    return None


def clean(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (dt.datetime, dt.date)):
        return f"{value.month}월 {value.day}일"
    s = str(value).replace("\r\n", "\n").replace(" ", " ").replace("﻿", "")
    lines = [ln.rstrip() for ln in s.split("\n")]
    return "\n".join(lines).strip()


def section_key(title: str) -> str:
    t = title.strip().strip("[]").strip().upper()
    for key in SECTION_ORDER:
        if t == key or t.replace(" ", "") == key.replace(" ", ""):
            return key
    if "DESIGN" in t:
        return "UX/UI DESIGN"
    return t


# ---------------------------------------------------------------------------
# 렌더링
# ---------------------------------------------------------------------------
LINK_MD = re.compile(r"\[([^\]\n]+)\]\((https?://[^)\s]+)\)")
URL_RE = re.compile(r"(?<![\"'=(])(https?://[^\s<]+)")


def rich(text: str) -> str:
    """셀 텍스트 → HTML. 줄바꿈은 <br>, [라벨](URL) 은 링크로."""
    text = clean(text)
    if not text:
        return "&nbsp;"
    out, pos = [], 0
    for m in LINK_MD.finditer(text):
        out.append(_plain_with_urls(text[pos:m.start()]))
        out.append(f'<a href="{html.escape(m.group(2), quote=True)}" target="_blank" '
                   f'style="color:#1155cc;">{html.escape(m.group(1))}</a>')
        pos = m.end()
    out.append(_plain_with_urls(text[pos:]))
    return "".join(out).replace("\n", "<br>")


def _plain_with_urls(s: str) -> str:
    parts, pos = [], 0
    for m in URL_RE.finditer(s):
        parts.append(html.escape(s[pos:m.start()]))
        url = m.group(1)
        parts.append(f'<a href="{html.escape(url, quote=True)}" target="_blank" '
                     f'style="color:#1155cc;">{html.escape(url)}</a>')
        pos = m.end()
    parts.append(html.escape(s[pos:]))
    return "".join(parts)


def td_style(field: str, *, header=False, last=False) -> str:
    st = ["border:1px solid #000000", "padding:2px 3px", "vertical-align:middle",
          "white-space:normal", "word-break:keep-all", "overflow-wrap:break-word",
          "font-size:10pt", f"font-family:{FONT}", "line-height:1.45"]
    if header:
        st += [f"background-color:{NOTE_HEADER_BG if last else HEADER_BG}",
               "color:#ffffff", "font-weight:bold", "text-align:center", "height:21px"]
    else:
        if field in CENTER:
            st.append("text-align:center")
        if field == "date":
            st.append("font-weight:bold")
    return ";".join(st) + ";"


def status_style(status: str) -> str:
    color = STATUS_COLORS.get(clean(status).replace("\n", " "))
    return f"color:{color};" if color else ""


def cell_html(field: str, value: str, *, rowspan: int = 1) -> str:
    style = td_style(field)
    if field == "status":
        style += status_style(value)
    content = rich(value)
    if field == "jira":
        keys = re.findall(r"\b([A-Z][A-Z0-9]+-\d+)\b", clean(value))
        if keys and not LINK_MD.search(clean(value)):
            content = "<br>".join(
                f'<a href="{JIRA_BASE}{k}" target="_blank" style="color:#1155cc;">{k}</a>'
                for k in keys)
    rs = f' rowspan="{rowspan}"' if rowspan > 1 else ""
    return f'<td{rs} style="{style}">{content}</td>'


# 메일에 표시하지 않는 열 (데이터에는 남겨 둔다). 설정 파일 "hide_columns" 로 바꿀 수 있다.
DEFAULT_HIDDEN = ["type"]   # 구분(PROJECT/TASK/DESIGN/DATA) 열은 메일에서 생략


def render_table(section: dict, hidden=()) -> str:
    key = section_key(section.get("title", ""))
    headers = section.get("headers") or DEFAULT_HEADERS.get(key) or DEFAULT_HEADERS["PROJECT"]
    keep = [i for i, f in enumerate(FIELDS) if f not in hidden]
    fields = [FIELDS[i] for i in keep]
    headers = [headers[i] if i < len(headers) else "" for i in keep]
    widths = [COL_WIDTHS[i] for i in keep]
    total = sum(widths)
    cols = "".join(f'<col width="{w}">' for w in widths)
    head = "".join(
        f'<td style="{td_style(fields[i], header=True, last=(i == len(headers) - 1))}">'
        f'{rich(h)}</td>' for i, h in enumerate(headers))
    rows_html = []
    for idx, row in enumerate(section.get("rows", []), 1):
        row = dict(row)
        row.setdefault("no", str(idx))
        steps = row.get("steps") or []
        if steps:
            n = len(steps)
            first = "".join(cell_html(f, row.get(f, ""), rowspan=n)
                            for f in fields if f not in STEP_FIELDS)
            for si, step in enumerate(steps):
                cells = "".join(cell_html(f, step.get(f, "")) for f in fields if f in STEP_FIELDS)
                rows_html.append(f"<tr>{first if si == 0 else ''}{cells}</tr>")
        else:
            rows_html.append("<tr>" + "".join(cell_html(f, row.get(f, "")) for f in fields) + "</tr>")
    return (f'<table cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;'
            f'table-layout:fixed;width:{total}px;font-size:10pt;font-family:{FONT};">'
            f"<colgroup>{cols}</colgroup><tbody><tr>{head}</tr>{''.join(rows_html)}</tbody></table>")


def p(inner: str, extra: str = "") -> str:
    return (f'<p style="margin:0;padding:0;line-height:20px;font-size:10pt;'
            f'font-family:{FONT};{extra}">{inner}</p>')


BLANK = p("&nbsp;")


def link_html(label: str, url: str, text: str | None = None) -> str:
    return (f'{html.escape(label)} <a href="{html.escape(url, quote=True)}" target="_blank" '
            f'style="color:#1155cc;">{html.escape(text or "링크")}</a>')


def render_body(rep: dict, cfg: dict) -> str:
    sender = cfg.get("sender", {})
    dept = rep.get("dept") or cfg.get("dept", "플랫폼기획부")
    name = sender.get("name", "")
    greeting = rep.get("greeting") or (f"안녕하세요, {dept} {name}입니다." if name else f"안녕하세요, {dept}입니다.")
    parts = [p(html.escape(greeting)), BLANK,
             p(f'<b>{html.escape(rep["period_label"])}</b> {html.escape(dept)} 주간보고 송부 드립니다.')]
    for note in rep.get("notes", []):
        parts.append(p(f"<i>{html.escape(note)}</i>", "color:#858585;"))
    parts.append(BLANK)

    links = rep.get("links", cfg.get("links", []))
    if links:
        parts.append(p("<b>■ 업무 현황/상세 업무</b>", "font-size:12pt;"))
        for i, ln in enumerate(links, 1):
            parts.append(p(f"{i}. " + link_html(ln["label"], ln["url"], ln.get("text"))))
        parts.append(BLANK)

    shared = rep.get("shared", cfg.get("shared", []))
    if shared:
        parts.append(p("<b>■ 공유 사항</b>", "font-size:12pt;"))
        for item in shared:
            if item.get("url"):
                parts.append(p(f'<a href="{html.escape(item["url"], quote=True)}" target="_blank" '
                               f'style="color:#1155cc;">{html.escape(item["label"])}</a>'))
            else:
                parts.append(p(rich(item["label"])))
        parts.append(BLANK)

    for sec in rep.get("sections", []):
        if not sec.get("rows"):
            continue
        parts += [BLANK, p(f'<b>[{html.escape(section_key(sec["title"]))}]</b>', "font-size:12pt;"), BLANK,
                  render_table(sec, cfg.get("hide_columns", DEFAULT_HIDDEN)), BLANK]

    closing = rep.get("closing") or cfg.get("closing") or ["이상입니다.", "감사합니다."]
    parts += [BLANK] + [p(html.escape(c)) for c in closing]
    sig = cfg.get("signature_html")
    if sig:
        parts += [BLANK, BLANK, sig]
    return f'<div style="font-family:{FONT};font-size:10pt;color:#000000;">' + "".join(parts) + "</div>"


def md_plain(s: str) -> str:
    return LINK_MD.sub(lambda m: f"{m.group(1)} ({m.group(2)})", s)


def render_text(rep: dict, cfg: dict) -> str:
    sender = cfg.get("sender", {})
    dept = rep.get("dept") or cfg.get("dept", "플랫폼기획부")
    name = sender.get("name", "")
    greeting = rep.get("greeting") or (f"안녕하세요, {dept} {name}입니다." if name else f"안녕하세요, {dept}입니다.")
    out = [greeting, "",
           f"{rep['period_label']} {dept} 주간보고 송부 드립니다."]
    out += rep.get("notes", [])
    out.append("")
    links = rep.get("links", cfg.get("links", []))
    if links:
        out.append("■ 업무 현황/상세 업무")
        out += [f"{i}. {ln['label']} : {ln['url']}" for i, ln in enumerate(links, 1)]
        out.append("")
    shared = rep.get("shared", cfg.get("shared", []))
    if shared:
        out.append("■ 공유 사항")
        out += [f"{s['label']}" + (f" : {s['url']}" if s.get("url") else "") for s in shared]
        out.append("")
    for sec in rep.get("sections", []):
        if not sec.get("rows"):
            continue
        out += ["", f"[{section_key(sec['title'])}]", ""]
        for idx, r in enumerate(sec["rows"], 1):
            head = " | ".join(x for x in [r.get("no") or str(idx),
                                          "" if "type" in cfg.get("hide_columns", DEFAULT_HIDDEN) else r.get("type", ""),
                                          r.get("jira", ""),
                                         r.get("platform", ""), clean(r.get("project", "")).replace("\n", " ")] if x)
            out.append(f"{head}")
            steps = r.get("steps") or [r]
            for s in steps:
                meta = " / ".join(x for x in [clean(s.get("status", "")), clean(s.get("date", "")),
                                              clean(s.get("owner", "")).replace("\n", " ")] if x)
                if meta:
                    out.append(f"  · 상태: {meta}")
                for label, f in (("상세", "detail"), ("일정", "schedule"), ("특이사항", "note")):
                    v = clean(s.get(f, ""))
                    if v:
                        out.append(f"  · {label}: " + md_plain(v).replace("\n", "\n    "))
            out.append("")
    closing = rep.get("closing") or cfg.get("closing") or ["이상입니다.", "감사합니다."]
    out += [""] + closing
    if cfg.get("signature_text"):
        out += ["", "", cfg["signature_text"]]
    return "\n".join(out).rstrip() + "\n"


PAGE = """<!doctype html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>주간보고 메일 {period}</title>
<style>
  body {{ margin:0; background:#f3f4f6; font-family:{font}; color:#111; }}
  .bar {{ position:sticky; top:0; z-index:2; background:#111827; color:#fff; padding:12px 16px;
          display:flex; flex-wrap:wrap; gap:8px; align-items:center; }}
  .bar button {{ font:inherit; font-size:13px; border:0; border-radius:6px; padding:8px 12px; cursor:pointer; }}
  .primary {{ background:#2563eb; color:#fff; font-weight:bold; }}
  .ghost {{ background:#374151; color:#fff; }}
  .meta {{ background:#fff; margin:16px; padding:12px 16px; border-radius:8px; font-size:13px; line-height:1.7; }}
  .meta b {{ display:inline-block; min-width:44px; }}
  .meta code {{ background:#f3f4f6; padding:1px 4px; border-radius:4px; }}
  #toast {{ margin-left:auto; font-size:13px; color:#a7f3d0; }}
  .paper {{ background:#fff; margin:0 16px 32px; padding:16px; overflow-x:auto; border-radius:8px; }}
</style></head>
<body>
<div class="bar">
  <button class="primary" onclick="copyBody()">본문 복사 (서식 포함)</button>
  <button class="ghost" onclick="copyText('subject')">제목 복사</button>
  {recipient_buttons}
  <span id="toast"></span>
</div>
<div class="meta">
  <div><b>제목</b> <span id="subject">{subject}</span></div>
  {recipient_rows}
  <div style="color:#6b7280;margin-top:6px">[본문 복사]를 누른 뒤 메일 작성 창 본문에 Ctrl+V(⌘V) 하세요. 표·색상·링크가 그대로 붙습니다.
  버튼이 동작하지 않으면 아래 흰 영역을 클릭 → Ctrl+A → Ctrl+C 로 복사해도 됩니다.</div>
</div>
<div class="paper" id="mail-body">{body}</div>
<script>
function toast(m){{const t=document.getElementById('toast');t.textContent=m;setTimeout(()=>t.textContent='',2500);}}
async function copyBody(){{
  const el=document.getElementById('mail-body');
  const htmlStr=el.innerHTML, text=el.innerText;
  try{{
    await navigator.clipboard.write([new ClipboardItem({{
      'text/html':new Blob([htmlStr],{{type:'text/html'}}),
      'text/plain':new Blob([text],{{type:'text/plain'}})}})]);
    toast('본문을 복사했습니다');
  }}catch(e){{
    const r=document.createRange();r.selectNodeContents(el);
    const s=window.getSelection();s.removeAllRanges();s.addRange(r);
    try{{document.execCommand('copy');toast('본문을 복사했습니다');}}catch(_){{toast('Ctrl+C 로 복사하세요');}}
  }}
}}
async function copyText(id){{
  const v=document.getElementById(id).innerText;
  try{{await navigator.clipboard.writeText(v);}}catch(e){{
    const ta=document.createElement('textarea');ta.value=v;document.body.appendChild(ta);ta.select();
    document.execCommand('copy');ta.remove();}}
  toast('복사했습니다');
}}
</script>
</body></html>
"""


def write_outputs(rep: dict, cfg: dict, outdir: Path) -> dict:
    outdir.mkdir(parents=True, exist_ok=True)
    dept = rep.get("dept") or cfg.get("dept", "플랫폼기획부")
    subject = rep.get("subject") or cfg.get("subject_format", "[{dept}] {period} 주간보고").format(
        dept=dept, period=rep["period_label"])
    body = render_body(rep, cfg)
    rec_buttons, rec_rows = [], []
    for key, label in (("to", "받는 사람"), ("cc", "참조")):
        vals = cfg.get("recipients", {}).get(key) or []
        if vals:
            rec_buttons.append(f'<button class="ghost" onclick="copyText(\'{key}\')">{label} 복사</button>')
            rec_rows.append(f'<div><b>{label}</b> <span id="{key}">{html.escape(", ".join(vals))}</span></div>')
    page = PAGE.format(period=html.escape(rep["period_label"]), font=FONT, subject=html.escape(subject),
                       recipient_buttons="".join(rec_buttons), recipient_rows="".join(rec_rows), body=body)
    files = {
        "mail.html": page,
        "body.html": f'<!doctype html><html><head><meta charset="utf-8"></head><body>{body}</body></html>',
        "mail.txt": render_text(rep, cfg),
        "subject.txt": subject + "\n",
        "report.json": json.dumps(rep, ensure_ascii=False, indent=2) + "\n",
    }
    for name, content in files.items():
        (outdir / name).write_text(content, encoding="utf-8")
    return {"subject": subject, "outdir": str(outdir),
            "rows": {section_key(s["title"]): len(s.get("rows", [])) for s in rep.get("sections", [])}}


# ---------------------------------------------------------------------------
# 기간 계산
# ---------------------------------------------------------------------------
def period_from_meeting(meeting: dt.date) -> str:
    """회의일(월요일) 기준 직전 주 월~금. 예) 9/28 회의 → '26.09.21 - 26.09.25'"""
    monday = meeting - dt.timedelta(days=meeting.weekday())
    start = monday - dt.timedelta(days=7)
    end = start + dt.timedelta(days=4)
    return f"{start:%y.%m.%d} - {end:%y.%m.%d}"


def meeting_date_from_sheet(name: str, today: dt.date | None = None) -> dt.date | None:
    # xlsx 로 내려받으면 탭 이름의 '/' 가 다른 문자로 바뀔 수 있어 구분자를 넓게 받는다
    m = re.search(r"(\d{1,2})\s*[/_.\-]\s*(\d{1,2})", name or "")
    if not m:
        return None
    today = today or dt.date.today()
    month, day = int(m.group(1)), int(m.group(2))
    year = today.year
    cand = dt.date(year, month, day)
    if (cand - today).days > 180:
        cand = dt.date(year - 1, month, day)
    return cand


# ---------------------------------------------------------------------------
# 표 → 행 데이터 (xlsx / eml 공통)
# ---------------------------------------------------------------------------
def grid_to_rows(header: list[str], body: list[list[dict]]) -> list[dict]:
    """header: 헤더 텍스트 목록. body: 행마다 [{text, group}] (group = 세로 병합 묶음 id).

    같은 group 으로 묶인 연속 행은 하나의 항목 + steps 로 합친다.
    """
    fields = [header_to_field(h) for h in header]
    rows, current, current_gid = [], None, None
    for line in body:
        rec = {}
        gid = None
        for f, cell in zip(fields, line):
            if not f:
                continue
            txt = clean(cell.get("text", ""))
            if cell.get("url") and txt and "[" not in txt and f not in ("jira",):
                txt = f"[{txt}]({cell['url']})"
            if f in rec and txt:
                rec[f] = (rec[f] + "\n" + txt).strip()
            else:
                rec.setdefault(f, txt)
            if f == "project" and cell.get("group") is not None:
                gid = cell["group"]
        if not any(rec.get(f) for f in FIELDS):
            continue
        step = {f: rec.get(f, "") for f in STEP_FIELDS}
        if gid is not None and gid == current_gid and current is not None:
            current["steps"].append(step)
            continue
        item = {f: rec.get(f, "") for f in FIELDS if f not in STEP_FIELDS}
        item.update(step)
        item["steps"] = [step]
        rows.append(item)
        current, current_gid = item, gid
    for r in rows:
        if len(r["steps"]) <= 1:
            r.pop("steps")
        else:
            for f in STEP_FIELDS:
                r.pop(f, None)
    return rows


# ---------------------------------------------------------------------------
# xlsx (구글 시트 '플랫폼기획부 주간 보고' 를 xlsx로 내려받은 파일)
# ---------------------------------------------------------------------------
def load_xlsx(path: Path, sheet: str | None) -> tuple[dict, str]:
    import openpyxl  # noqa: WPS433

    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet:
        if sheet not in wb.sheetnames:
            key = re.sub(r"[^0-9A-Za-z가-힣]", "", sheet)
            cand = [s for s in wb.sheetnames if key and key in re.sub(r"[^0-9A-Za-z가-힣]", "", s)]
            if not cand:
                raise SystemExit(f"시트 '{sheet}' 를 찾지 못했습니다. 시트 목록: {wb.sheetnames}")
            sheet = cand[0]
        ws = wb[sheet]
    else:
        weekly = [s for s in wb.sheetnames if re.match(r"^주간\s*보고\s*\(\d{1,2}\D\d{1,2}\)$", s.strip())]
        if not weekly:
            raise SystemExit(f"'주간 보고(M/D)' 형식의 탭이 없습니다. --sheet 로 지정하세요. 시트 목록: {wb.sheetnames}")
        weekly.sort(key=lambda s: meeting_date_from_sheet(s) or dt.date.min)
        sheet = weekly[-1]
        ws = wb[sheet]

    merged_top, rowspans = {}, {}
    for rng in ws.merged_cells.ranges:
        rowspans[(rng.min_row, rng.min_col)] = rng.max_row - rng.min_row + 1
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merged_top[(r, c)] = (rng.min_row, rng.min_col)

    def cell(r, c):
        tr, tc = merged_top.get((r, c), (r, c))
        ce = ws.cell(tr, tc)
        url = ce.hyperlink.target if ce.hyperlink is not None else None
        return {"text": clean(ce.value) if (tr, tc) == (r, c) or tr == r else "",
                "url": url if tr == r else None,
                "group": (tr, tc) if rowspans.get((tr, tc), 1) > 1 else None}

    def row_texts(r):
        return [clean(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]

    sections, title, i = [], None, 1
    intro_lines = []
    while i <= ws.max_row:
        texts = row_texts(i)
        nonempty = [t for t in texts if t]
        joined = " ".join(nonempty)
        if len(nonempty) == 1 and re.fullmatch(r"\[[^\]]+\]", nonempty[0]):
            title = nonempty[0]
            i += 1
            continue
        if any(norm_header(t) == "jirakey" for t in texts):
            cols = [c for c in range(1, ws.max_column + 1) if row_texts(i)[c - 1]]
            c0, c1 = cols[0], cols[-1]
            header = [clean(ws.cell(i, c).value) for c in range(c0, c1 + 1)]
            body, j = [], i + 1
            while j <= ws.max_row:
                t = row_texts(j)
                ne = [x for x in t if x]
                if len(ne) == 1 and re.fullmatch(r"\[[^\]]+\]", ne[0]):
                    break
                if any(norm_header(x) == "jirakey" for x in t):
                    break
                if not ne and not any((j, c) in merged_top for c in range(c0, c1 + 1)):
                    # 빈 줄 2개 연속이면 표 끝
                    if j + 1 > ws.max_row or not any(row_texts(j + 1)):
                        break
                    j += 1
                    continue
                body.append([cell(j, c) for c in range(c0, c1 + 1)])
                j += 1
            sections.append({"title": section_key(title or f"SECTION{len(sections) + 1}"),
                             "rows": grid_to_rows(header, body)})
            title = None
            i = j
            continue
        if not sections and joined:
            intro_lines.append(joined)
        i += 1
    return {"sections": sections, "intro_lines": intro_lines}, sheet


# ---------------------------------------------------------------------------
# eml (지난주 발송 메일)
# ---------------------------------------------------------------------------
class _MailParser(HTMLParser):
    """메일 HTML 에서 일반 텍스트 줄, 링크, 데이터 표를 순서대로 뽑는다.

    메일 프로그램이 본문 전체를 감싸는 레이아웃 표(id=contentT 등)는 투명하게 취급하고,
    구글 시트에서 붙여넣은 표(data-sheets-root) 또는 그 밖의 표만 데이터 표로 본다.
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.events = []          # ("text", str) | ("link", {...}) | ("table", rows)
        self.kinds = []           # 열린 table 들의 종류: layout / data / hidden
        self.table = None         # 현재 데이터 표
        self.row = None
        self.cell = None
        self.anchor = None
        self.buf = []

    def _hidden(self):
        return "hidden" in self.kinds

    def _flush_text(self):
        t = "".join(self.buf)
        self.buf = []
        for line in t.split("\n"):
            line = line.replace("\u00a0", " ").strip()
            if line:
                self.events.append(("text", line))

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "table":
            style = (a.get("style") or "").replace(" ", "").lower()
            if "display:none" in style:
                kind = "hidden"
            elif a.get("id") == "contentT" or (self.table is None and "data-sheets-root" not in a
                                                and a.get("role") == "presentation"):
                kind = "layout"
            elif self.table is None:
                kind = "data"
                self._flush_text()
                self.table = []
            else:
                kind = "layout"   # 데이터 표 안의 중첩 표는 셀 텍스트로 취급
            self.kinds.append(kind)
            return
        if self._hidden():
            return
        if self.table is not None and "data" in self.kinds:
            if tag == "tr":
                self.row = []
            elif tag in ("td", "th"):
                self.cell = {"text": "", "rowspan": int(a.get("rowspan") or 1),
                             "colspan": int(a.get("colspan") or 1), "url": None}
            elif self.cell is not None:
                if tag == "br" or (tag in ("p", "div") and self.cell["text"]):
                    self.cell["text"] += "\n"
                elif tag == "a":
                    self.cell["url"] = self.cell["url"] or a.get("href")
            return
        if tag in ("p", "div", "br", "tr", "li"):
            self.buf.append("\n")
        if tag == "a":
            line = "".join(self.buf).split("\n")[-1].replace("\u00a0", " ").strip()
            self.anchor = {"href": a.get("href", ""), "text": "", "line": line}

    def handle_endtag(self, tag):
        if tag == "table" and self.kinds:
            kind = self.kinds.pop()
            if kind == "data" and self.table is not None:
                self.events.append(("table", self.table))
                self.table = None
            return
        if self._hidden():
            return
        if self.table is not None and "data" in self.kinds:
            if tag in ("td", "th") and self.cell is not None and self.row is not None:
                self.row.append(self.cell)
                self.cell = None
            elif tag == "tr" and self.row is not None:
                self.table.append(self.row)
                self.row = None
            return
        if tag == "a" and self.anchor is not None:
            self.events.append(("link", self.anchor))
            self.anchor = None
        if tag in ("p", "div", "li"):
            self.buf.append("\n")
            self._flush_text()

    def handle_data(self, data):
        if self._hidden():
            return
        if self.table is not None and "data" in self.kinds:
            if self.cell is not None:
                self.cell["text"] += data
            return
        if self.anchor is not None:
            self.anchor["text"] += data
        self.buf.append(data)


def expand_rowspans(rows: list[list[dict]]) -> list[list[dict]]:
    """rowspan 을 풀어서 각 행을 같은 열 수로 맞추고, 병합 셀에는 group id 를 단다."""
    grid, pending = [], {}
    gid = 0
    for r_i, row in enumerate(rows):
        out, col, it = [], 0, iter(row)
        while True:
            if col in pending and pending[col][1] > 0:
                ref, left = pending[col]
                out.append(ref)
                pending[col] = (ref, left - 1)
                col += 1
                continue
            try:
                c = next(it)
            except StopIteration:
                break
            entry = {"text": c["text"], "url": c.get("url"), "group": None}
            if c["rowspan"] > 1:
                gid += 1
                entry["group"] = gid
                pending[col] = (entry, c["rowspan"] - 1)
            for k in range(c.get("colspan", 1)):
                out.append(entry if k == 0 else {"text": "", "url": None, "group": entry["group"]})
                col += 1
        while col in pending and pending[col][1] > 0:
            ref, left = pending[col]
            out.append(ref)
            pending[col] = (ref, left - 1)
            col += 1
        grid.append(out)
    return grid


def load_eml(path: Path) -> dict:
    msg = email.message_from_bytes(path.read_bytes(), policy=policy.default)
    body_html = None
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            body_html = part.get_content()
            break
    if body_html is None:
        raise SystemExit("메일에 HTML 본문이 없습니다.")
    mp = _MailParser()
    mp.feed(body_html)
    mp._flush_text()

    rep = {"subject_prev": str(msg["subject"] or ""), "notes": [], "links": [], "shared": [], "sections": []}
    m = re.search(r"(\d{2}\.\d{2}\.\d{2})\s*-\s*(\d{2}\.\d{2}\.\d{2})", rep["subject_prev"])
    rep["period_label"] = f"{m.group(1)} - {m.group(2)}" if m else ""

    mode, title, last_text = None, None, ""
    sections = []
    for kind, val in mp.events:
        if kind == "text":
            if val.startswith("■") and "업무" in val:
                mode = "links"
            elif val.startswith("■") and "공유" in val:
                mode = "shared"
            elif re.fullmatch(r"\[[^\]]+\]", val):
                title, mode = val, "sections"
            elif val.startswith("*") and not sections:
                rep["notes"].append(val)
            last_text = val
        elif kind == "link":
            if mode == "links":
                label = re.sub(r"^\d+\.\s*", "", val.get("line") or last_text).strip()
                rep["links"].append({"label": label, "url": val["href"], "text": val["text"].strip() or "링크"})
            elif mode == "shared":
                rep["shared"].append({"label": val["text"].strip(), "url": val["href"]})
        elif kind == "table":
            if not val:
                continue
            header = [clean(c["text"]) for c in val[0]]
            has_header = any(norm_header(h) == "jirakey" for h in header)
            if has_header:
                grid = expand_rowspans(val[1:])
                sections.append({"title": section_key(title or ""), "headers": header,
                                 "rows": grid_to_rows(header, grid)})
                title = None
            elif sections and val and max(len(r) for r in val) >= 9:
                # 헤더 없이 바로 이어 붙은 표 (예: PROJECT 아래 '홀딩' 행)
                prev = sections[-1]
                grid = expand_rowspans(val)
                hdr = prev["headers"]
                extra = grid_to_rows(hdr[1:] if len(grid[0]) == len(hdr) - 1 else hdr, grid)
                for r in extra:
                    r["no"] = ""
                prev["rows"].extend(extra)
    for s in sections:
        s["headers"] = [h.replace("특이 사항(외부 공유용)", "특이 사항\n(외부 공유용)") for h in s["headers"]]
    rep["sections"] = sections
    return rep


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def load_config(path: str | None) -> dict:
    cands = [Path(path)] if path else [HERE / "config.local.json", HERE / "config.example.json"]
    for c in cands:
        if c.exists():
            return json.loads(c.read_text(encoding="utf-8"))
    return {}


def finalize(rep: dict, cfg: dict, args) -> dict:
    if getattr(args, "period", None):
        rep["period_label"] = args.period
    if not rep.get("period_label"):
        raise SystemExit("보고 기간을 알 수 없습니다. --period '26.11.02 - 26.11.06' 처럼 지정하세요.")
    for note in getattr(args, "note", None) or []:
        rep.setdefault("notes", []).append(note)
    for key in ("links", "shared"):
        if not rep.get(key) and cfg.get(key):
            rep[key] = cfg[key]
    order = {k: i for i, k in enumerate(SECTION_ORDER)}
    rep["sections"] = sorted(rep.get("sections", []),
                             key=lambda s: order.get(section_key(s.get("title", "")), 99))
    return rep


def main(argv=None):
    ap = argparse.ArgumentParser(description="플랫폼기획부 주간보고 메일 생성기")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(sp, out_required=True):
        sp.add_argument("-o", "--out", required=out_required, help="출력 폴더 (from-eml 은 .json 파일)")
        sp.add_argument("--config", help="설정 파일 (기본: config.local.json → config.example.json)")
        sp.add_argument("--period", help="보고 기간 표기. 예: '26.11.02 - 26.11.06'")
        sp.add_argument("--note", action="append", help="기간 아래 회색 메모. 예: '*24~25 추석 연휴'")

    b = sub.add_parser("build", help="JSON → 메일")
    b.add_argument("json")
    common(b)

    x = sub.add_parser("from-xlsx", help="회의록 xlsx → 메일")
    x.add_argument("xlsx")
    x.add_argument("--sheet", help="탭 이름. 생략 시 가장 최근 '주간 보고(M/D)' 탭")
    x.add_argument("--json-only", action="store_true", help="메일은 만들지 않고 report.json 만 저장")
    common(x)

    e = sub.add_parser("from-eml", help="지난 메일(.eml) → JSON (+ 선택적으로 메일 재생성)")
    e.add_argument("eml")
    e.add_argument("--render", metavar="DIR", help="같은 내용으로 메일도 생성할 폴더")
    common(e)

    args = ap.parse_args(argv)
    cfg = load_config(args.config)

    if args.cmd == "build":
        rep = json.loads(Path(args.json).read_text(encoding="utf-8"))
        rep = finalize(rep, cfg, args)
        info = write_outputs(rep, cfg, Path(args.out))
    elif args.cmd == "from-xlsx":
        rep, sheet = load_xlsx(Path(args.xlsx), args.sheet)
        rep["source"] = {"file": Path(args.xlsx).name, "sheet": sheet}
        if not args.period:
            md = meeting_date_from_sheet(sheet)
            if md:
                rep["period_label"] = period_from_meeting(md)
                rep["period_inferred_from"] = f"탭 '{sheet}' 회의일 {md:%Y-%m-%d} 기준 직전 주 월~금"
        rep = finalize(rep, cfg, args)
        if args.json_only:
            out = Path(args.out)
            out.mkdir(parents=True, exist_ok=True)
            (out / "report.json").write_text(json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")
            info = {"outdir": str(out), "rows": {s["title"]: len(s["rows"]) for s in rep["sections"]}}
        else:
            info = write_outputs(rep, cfg, Path(args.out))
        info["sheet"] = sheet
        info["period"] = rep["period_label"]
    else:
        rep = load_eml(Path(args.eml))
        rep = finalize(rep, cfg, args)
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(rep, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        info = {"json": str(out), "rows": {s["title"]: len(s["rows"]) for s in rep["sections"]}}
        if args.render:
            info["render"] = write_outputs(rep, cfg, Path(args.render))
    print(json.dumps(info, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    sys.exit(main())
