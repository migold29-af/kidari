#!/usr/bin/env python3
"""월간 지표 리포트 요약 메일 생성기 (표준 라이브러리만 사용).

content.json(요약 내용) + config.json(수신자·고정 문구)을 받아
  - mail.html : 브라우저로 열어 전체 복사 → 사내 웹메일 에디터에 붙여넣기용
  - mail.eml  : 메일 클라이언트(Outlook 등)에서 열면 '보내기 전 초안'으로 열림 (X-Unsent: 1)
  - mail.txt  : 텍스트 버전
을 만든다.

사용법:
  python3 build_mail.py <content.json> [--config config.json] [--out 출력폴더] [--attach 리포트.pdf]

인라인 마크업 (content.json 문자열 안에서):
  **텍스트**  → 굵게
  __텍스트__  → 굵게 + 밑줄 (원본 메일의 핵심 문장 강조 스타일)
"""
from __future__ import annotations

import argparse
import html
import json
import mimetypes
import os
import re
import sys
import uuid
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

HERE = Path(__file__).resolve().parent

FONT = "font-family:'맑은 고딕','Malgun Gothic',돋움,dotum,AppleGothic,sans-serif;"
BASE = FONT + "font-size:10pt;line-height:20px;color:#222;"
RULE = "━━━━━━━━━━━━━━━━━━━━━━━━"
C_PREV = "#b0b0b0"   # 전월 막대 (원본 차트 회색)
C_CURR = "#4a90d9"   # 금월 막대 (원본 차트 파랑)
C_TH = "#6aa0f0"     # 표 헤더
C_BOX = "#eaf2fd"    # 섹션 요약 박스
C_CALLOUT = "#fde7b0"  # 하이라이트 문장
ROW_BG = {"down": "#fdecec", "up": "#eaf2fd", "neutral": "#f2f3f5"}

DELTA_RE = re.compile(r"^[+\-−]\s?\d[\d,]*(\.\d+)?\s?(%p|%|p)?$")


# ---------------------------------------------------------------- helpers
def inline(text: str) -> str:
    """HTML escape 후 **굵게**, __굵게+밑줄__ 변환."""
    s = html.escape(str(text))
    s = re.sub(r"__(.+?)__", r"<b><u>\1</u></b>", s)
    s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
    return s


def plain(text: str) -> str:
    return re.sub(r"(\*\*|__)(.+?)\1", r"\2", str(text))


def with_arrow(cell: str) -> str:
    """'-2.85%' → '-2.85% ↓', '+7.69%' → '+7.69% ↑' (이미 화살표가 있으면 그대로)."""
    c = str(cell).strip()
    if "↑" in c or "↓" in c or not DELTA_RE.match(c):
        return c
    return f"{c} ↓" if c[0] in "-−" else f"{c} ↑"


def p(content: str, extra: str = "") -> str:
    return f'<p style="{BASE}margin:0;padding:0;{extra}">{content}</p>'


def spacer(n: int = 1) -> str:
    return p("&nbsp;") * n


def ul(items: list[str]) -> str:
    lis = "".join(f'<li style="{BASE}margin:0 0 4px 0;">{inline(i)}</li>' for i in items)
    return f'<ul style="{BASE}margin:4px 0 4px 0;padding-left:22px;">{lis}</ul>'


def header_block(title: str) -> str:
    return p(f"{RULE}<br><b>{html.escape(title)}</b><br>{RULE}", "line-height:150%;")


# ---------------------------------------------------------------- tables
def render_table(t: dict) -> str:
    cols = t["columns"]
    groups = t.get("header_groups")  # [[label, span], ...]
    td = f"{BASE}border:1px solid #d9dde3;padding:6px 10px;"
    th = f"{td}background:{C_TH};color:#fff;font-weight:normal;text-align:center;"
    out = [f'<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:6px 0;">']
    if t.get("title"):
        out.append(f'<caption style="{BASE}text-align:left;font-weight:bold;padding:4px 0;">{inline(t["title"])}</caption>')
    if groups:
        row1, row2, i = [], [], 0
        for label, span in groups:
            if span == 1:
                row1.append(f'<th rowspan="2" style="{th}">{inline(label)}</th>')
            else:
                row1.append(f'<th colspan="{span}" style="{th}">{inline(label)}</th>')
                row2 += [f'<th style="{th}">{inline(c)}</th>' for c in cols[i:i + span]]
            i += span
        out.append("<tr>" + "".join(row1) + "</tr><tr>" + "".join(row2) + "</tr>")
    else:
        out.append("<tr>" + "".join(f'<th style="{th}">{inline(c)}</th>' for c in cols) + "</tr>")
    for row in t["rows"]:
        cells, hl = (row["cells"], row.get("highlight")) if isinstance(row, dict) else (row, None)
        bg = f"background:{ROW_BG[hl]};" if hl else ""
        bold = "font-weight:bold;" if hl == "up" else ""
        tds = []
        for j, c in enumerate(cells):
            align = "text-align:left;" if j == 0 and not t.get("first_col_center") else "text-align:center;"
            tds.append(f'<td style="{td}{align}{bg}{bold}">{inline(with_arrow(c))}</td>')
        out.append("<tr>" + "".join(tds) + "</tr>")
    out.append("</table>")
    return "".join(out)


def table_text(t: dict) -> str:
    lines = []
    if t.get("title"):
        lines.append(f"[{plain(t['title'])}]")
    lines.append(" | ".join(plain(c) for c in t["columns"]))
    for row in t["rows"]:
        cells = row["cells"] if isinstance(row, dict) else row
        lines.append(" | ".join(plain(with_arrow(c)) for c in cells))
    return "\n".join(lines)


# ---------------------------------------------------------------- charts
def fmt_num(v: float, unit: str) -> str:
    if unit == "%":
        return f"{v:.1f}%"
    return f"{v:,.0f}"


def render_chart(c: dict) -> str:
    """메일 클라이언트에서도 깨지지 않는 표 기반 가로 막대 차트 (전월=회색, 금월=파랑)."""
    unit = c.get("unit", "")
    show = c.get("show_values", "both")  # both | curr | none (역산·근사치인 전월 값은 숨김)
    labels, prev, curr = c["labels"], c["prev"], c["curr"]
    vmax = max([v for v in prev + curr if v is not None] or [1]) or 1
    bar_w = 300
    cell = f"{BASE}padding:1px 6px;"
    prev_lbl = html.escape(c.get("prev_label", "전월"))
    curr_lbl = html.escape(c.get("curr_label", "금월"))
    rows = []
    for lab, pv, cv in zip(labels, prev, curr):
        for k, (v, color) in enumerate(((pv, C_PREV), (cv, C_CURR))):
            w = 0 if v is None else max(1, round(bar_w * v / vmax))
            label_td = (f'<td rowspan="2" style="{cell}text-align:right;white-space:nowrap;'
                        f'border-bottom:1px solid #eee;">{html.escape(lab)}</td>') if k == 0 else ""
            val = "-" if v is None else fmt_num(v, unit)
            if show == "none" or (show == "curr" and k == 0):
                val = ""
            rows.append(
                f"<tr>{label_td}<td style=\"{cell}\">"
                f'<table cellpadding="0" cellspacing="0"><tr>'
                f'<td width="{w}" height="12" bgcolor="{color}" style="background:{color};width:{w}px;height:12px;font-size:1px;line-height:1px;">&nbsp;</td>'
                f'<td style="{FONT}font-size:8pt;color:#555;padding-left:6px;white-space:nowrap;">{val}</td>'
                f"</tr></table></td></tr>")
    legend = (f'<span style="color:{C_PREV};">■</span> {prev_lbl} &nbsp; '
              f'<span style="color:{C_CURR};">■</span> {curr_lbl}')
    return (f'<table cellpadding="0" cellspacing="0" style="margin:8px 0;border:1px solid #e3e6ea;padding:8px;">'
            f'<tr><td colspan="2" style="{BASE}font-weight:bold;text-align:center;padding-bottom:4px;">{html.escape(c.get("title", ""))}</td></tr>'
            f'<tr><td colspan="2" style="{FONT}font-size:8pt;text-align:right;padding-bottom:4px;">{legend}</td></tr>'
            + "".join(rows) + "</table>")


# ---------------------------------------------------------------- images
def render_image(img: dict, base_dir: Path, inline_imgs: list) -> str:
    path = (base_dir / img["path"]).resolve()
    cid = f"{uuid.uuid4()}@kpi-mail"
    inline_imgs.append((cid, path))
    width = img.get("width", 720)
    # mail.html 미리보기에서는 파일 경로, eml에서는 cid로 치환
    return f'<img data-cid="{cid}" src="{path.as_uri()}" style="width:{width}px;max-width:100%;" alt="{html.escape(img.get("alt", ""))}">'


# ---------------------------------------------------------------- blocks
def render_box(bullets: list[str], callout: str | None) -> str:
    inner = ul(bullets) if bullets else ""
    if callout:
        inner += p(f'<span style="color:#e0457b;">|</span> '
                   f'<span style="background:{C_CALLOUT};">{inline(callout)}</span>', "margin-top:6px;")
    return (f'<table cellpadding="0" cellspacing="0" width="760" style="max-width:100%;margin:6px 0 10px 0;">'
            f'<tr><td style="{BASE}background:{C_BOX};padding:10px 14px;">{inner}</td></tr></table>')


def render_notice(text: str) -> str:
    """리포트 상단 '📢 안내사항' 박스 (예: 대시보드 종료 공지)."""
    return (f'<table cellpadding="0" cellspacing="0" width="760" style="max-width:100%;margin:10px 0;">'
            f'<tr><td style="{BASE}background:#fff4e5;border-left:4px solid #f5a623;padding:10px 14px;">'
            f"<b>📢 안내사항</b><br>{inline(text)}</td></tr></table>")


def render_section(sec: dict, base_dir: Path, imgs: list) -> str:
    h = [p(f'<b style="font-size:11pt;">{inline(sec["title"])}</b>', "margin-top:4px;")]
    if sec.get("bullets") or sec.get("callout"):
        h.append(render_box(sec.get("bullets", []), sec.get("callout")))
    for ch in sec.get("charts", []):
        h.append(render_chart(ch))
    for t in sec.get("tables", []):
        h.append(render_table(t))
    for im in sec.get("images", []):
        h.append(p(render_image(im, base_dir, imgs)))
    for n in sec.get("notes", []):
        h.append(p(f"<b>{inline(n['label'])}</b>", "margin-top:10px;"))
        h.append(p(inline(n["text"])))
    h.append(spacer(2))
    return "".join(h)


def section_text(sec: dict) -> str:
    out = [plain(sec["title"])]
    out += [f"- {plain(b)}" for b in sec.get("bullets", [])]
    if sec.get("callout"):
        out.append(f"| {plain(sec['callout'])}")
    for t in sec.get("tables", []):
        out.append(table_text(t))
    for n in sec.get("notes", []):
        out.append(f"{plain(n['label'])}: {plain(n['text'])}")
    return "\n".join(out)


# ---------------------------------------------------------------- build
def build(content: dict, cfg: dict, base_dir: Path, attach: Path | None):
    m, pm = content["month"], content.get("prev_month", (content["month"] - 2) % 12 + 1)
    yy = content["year"] % 100
    svc = cfg["service_name"]
    subject = content.get("subject") or cfg["subject_format"].format(yy=yy, m=m, service=svc)
    imgs: list = []
    attach_name = attach.name if attach else content.get("attachment_name")

    # ---- 인트로
    b = [p(f"안녕하세요,<br>{html.escape(cfg['team_name'])} 입니다.<br><br>"
           f'<b><span style="color:#ff0000;">{m}월</span> '
           f"{html.escape(cfg['intro_line'].format(service=svc))}</b>")]
    b.append(p(html.escape(cfg["intro_note"]).replace("\n", "<br>")))
    links = []
    # content에 dashboard_url 키가 있으면 그 값(빈 문자열 = 링크 생략), 없으면 config 값
    dash = content["dashboard_url"] if "dashboard_url" in content else cfg.get("dashboard_url")
    if dash:
        sub = ""
        if attach_name:
            sub = (f'<br><i style="color:#404040;">* 링크 접속이 어려우신 경우 첨부된 PDF 파일('
                   f"{html.escape(attach_name)})을 확인해 주시기 바랍니다.</i>")
        links.append(f'대시보드 링크: <a href="{html.escape(dash)}" style="text-decoration:none;">{html.escape(dash)}</a>{sub}')
    if content.get("confluence_url"):
        title = content.get("confluence_title") or f"{m}월 지표 분석 리포트"
        links.append(f'<span style="color:#404040;">컨플루언스 링크 : <a href="{html.escape(content["confluence_url"])}">{html.escape(title)}</a></span>')
    if links:
        b.append(f'<ul style="{BASE}margin:4px 0;padding-left:22px;">'
                 + "".join(f'<li style="{BASE}">{l}</li>' for l in links) + "</ul>")
    if content.get("notice"):
        b.append(render_notice(content["notice"]))
    b.append(spacer(2))

    # ---- 핵심 요약
    s = content["summary"]
    b.append(header_block("✅ 핵심 요약"))
    b.append(spacer())
    b.append(p(f'<b style="font-size:11pt;">{inline(s["headline"])}</b>'))
    b.append(ul(s.get("bullets", [])))
    for ch in s.get("charts", []):
        b.append(render_chart(ch))
    for t in s.get("tables", []):
        b.append(render_table(t))
    for im in s.get("images", []):
        b.append(p(render_image(im, base_dir, imgs)))
    b.append(spacer(3))

    # ---- 주요 지표
    b.append(header_block("📊 주요 지표"))
    b.append(spacer())
    for sec in content.get("sections", []):
        b.append(render_section(sec, base_dir, imgs))

    # ---- 맺음말
    if content.get("footnote"):
        b.append(p(f'<i style="color:#666;">{inline(content["footnote"])}</i>'))
    b.append(p(f"{RULE}<br><br>"))
    b.append(p(html.escape(cfg["closing_note"]).replace("\n", "<br>")))
    if cfg.get("analyst"):
        b.append(f'<ul style="{BASE}margin:4px 0;padding-left:22px;"><li style="{BASE}"><b>'
                 f"데이터분석 담당자 : {html.escape(cfg['analyst'])}</b></li></ul>")
    b.append(spacer())
    b.append(p("감사합니다."))
    b.append(p(f"{html.escape(cfg['team_name'])} 드림."))
    if cfg.get("signature_html"):
        b.append(spacer(2) + cfg["signature_html"])

    body_html = f'<div style="{BASE}">' + "".join(b) + "</div>"

    # ---- 텍스트 버전
    t = [f"안녕하세요,\n{cfg['team_name']} 입니다.\n",
         f"{m}월 {cfg['intro_line'].format(service=svc)}\n", cfg["intro_note"], ""]
    if dash:
        t.append(f"대시보드 링크: {dash}")
    if content.get("confluence_url"):
        t.append(f"컨플루언스 링크: {content['confluence_url']}")
    if content.get("notice"):
        t += ["", f"📢 안내사항: {plain(content['notice'])}"]
    t += ["", RULE, "✅ 핵심 요약", RULE, plain(s["headline"])]
    t += [f"- {plain(x)}" for x in s.get("bullets", [])]
    t += [table_text(x) for x in s.get("tables", [])]
    t += ["", RULE, "📊 주요 지표", RULE]
    t += [section_text(sec) + "\n" for sec in content.get("sections", [])]
    if content.get("footnote"):
        t.append(plain(content["footnote"]))
    t += [RULE, "", cfg["closing_note"]]
    if cfg.get("analyst"):
        t.append(f"- 데이터분석 담당자 : {cfg['analyst']}")
    t += ["", "감사합니다.", f"{cfg['team_name']} 드림."]
    body_txt = "\n".join(t)

    return subject, body_html, body_txt, imgs, pm


COPY_PAGE = r"""<title>__TITLE__</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap">
<style>
/* layout: 상단 = 제목·수신자·복사 버튼 / 하단 = 메일 본문(흰 종이, 메일에 들어갈 모습 그대로) */
:root{--bg:#f3f5f8;--panel:#ffffff;--fg:#1d2330;--muted:#5d6676;--line:#dde2ea;--accent:#2f6fdb;--accent-fg:#ffffff;--ok:#1f8a4c;
--ui:'Noto Sans KR','맑은 고딕','Malgun Gothic',system-ui,sans-serif}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#151a22;--panel:#1e2530;--fg:#e6eaf0;--muted:#9aa4b4;--line:#323b49;--accent:#5b8ff0;--accent-fg:#0d1220;--ok:#4cc584;color-scheme:dark}}
:root[data-theme="dark"]{--bg:#151a22;--panel:#1e2530;--fg:#e6eaf0;--muted:#9aa4b4;--line:#323b49;--accent:#5b8ff0;--accent-fg:#0d1220;--ok:#4cc584;color-scheme:dark}
body{background:var(--bg);color:var(--fg);font-family:var(--ui);font-size:14px;padding-block:20px 40px;
padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
.wrap{max-width:900px;margin:0 auto;display:grid;gap:16px}
h1{font-size:18px;font-weight:700;margin:0;text-wrap:balance}
.lead{color:var(--muted);margin:4px 0 0;line-height:1.6}
.meta{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:6px 16px}
.row{display:grid;grid-template-columns:72px minmax(0,1fr) auto;gap:12px;align-items:start;padding:10px 0;border-bottom:1px solid var(--line)}
.row:last-child{border-bottom:0}
.row dt{color:var(--muted);font-size:12px;letter-spacing:.04em;padding-top:2px}
.row dd{margin:0;line-height:1.6;word-break:break-all}
button{font:inherit;cursor:pointer;border-radius:8px;border:1px solid var(--line);background:var(--panel);color:var(--fg);padding:5px 12px;font-size:13px;white-space:nowrap}
button:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.primary{background:var(--accent);color:var(--accent-fg);border-color:var(--accent);font-weight:700;font-size:15px;padding:12px 20px}
.bar{display:flex;flex-wrap:wrap;gap:12px;align-items:center}
.status{color:var(--ok);font-weight:500;min-height:1.4em}
.paper-scroll{overflow-x:auto;border:1px solid var(--line);border-radius:10px}
.paper{background:#ffffff;color:#222222;padding:24px;min-width:0}
@media (max-width:520px){.row{grid-template-columns:minmax(0,1fr) auto}.row dt{grid-column:1/-1}}
</style>
<div class="wrap">
  <header>
    <h1>__TITLE__</h1>
    <p class="lead">① <b>본문 복사</b>를 누르고 사내 메일 에디터에 붙여넣기(Ctrl+V) ② 제목·받는 사람·참조는 옆의 복사 버튼으로 넣기__ATTACH_HINT__</p>
  </header>
  <dl class="meta">
    <div class="row"><dt>제목</dt><dd id="v-subject">__SUBJECT__</dd><button data-copy="v-subject">복사</button></div>
    <div class="row"><dt>받는 사람</dt><dd id="v-to">__TO__</dd><button data-copy="v-to">복사</button></div>
    <div class="row"><dt>참조</dt><dd id="v-cc">__CC__</dd><button data-copy="v-cc">복사</button></div>
  </dl>
  <div class="bar"><button class="primary" id="copy-body">본문 복사</button><span class="status" id="status" role="status"></span></div>
  <div class="paper-scroll"><div class="paper" id="mail-body">__BODY__</div></div>
</div>
<script>
const statusEl=document.getElementById('status');
function say(t){statusEl.textContent=t;clearTimeout(say.t);say.t=setTimeout(()=>statusEl.textContent='',4000)}
function selectEl(el){const r=document.createRange();r.selectNodeContents(el);const s=getSelection();s.removeAllRanges();s.addRange(r)}
function legacyCopy(el){selectEl(el);let ok=false;try{ok=document.execCommand('copy')}catch(e){}return ok}
document.getElementById('copy-body').addEventListener('click',async()=>{
  const el=document.getElementById('mail-body');
  try{
    await navigator.clipboard.write([new ClipboardItem({
      'text/html':new Blob([el.innerHTML],{type:'text/html'}),
      'text/plain':new Blob([el.innerText],{type:'text/plain'})})]);
    say('본문을 복사했습니다. 메일 에디터에 붙여넣으세요.');
  }catch(e){
    if(legacyCopy(el)){getSelection().removeAllRanges();say('본문을 복사했습니다. 메일 에디터에 붙여넣으세요.')}
    else say('본문을 선택해 두었습니다. Ctrl+C(⌘C)로 복사하세요.');
  }
});
document.querySelectorAll('[data-copy]').forEach(b=>b.addEventListener('click',async()=>{
  const el=document.getElementById(b.dataset.copy);
  try{await navigator.clipboard.writeText(el.innerText);say('복사했습니다.')}
  catch(e){if(legacyCopy(el)){getSelection().removeAllRanges();say('복사했습니다.')}else say('선택해 두었습니다. Ctrl+C(⌘C)로 복사하세요.')}
}));
</script>
"""


def write_copy_page(out: Path, subject: str, body_html: str, cfg: dict, attach: Path | None, m: int):
    """웹에서 열어 '본문 복사' 버튼으로 서식째 복사하는 페이지 (Artifact 게시용)."""
    hint = f" ③ 첨부: {html.escape(attach.name)}" if attach else ""
    page = (COPY_PAGE.replace("__TITLE__", f"{m}월 지표 리포트 메일")
            .replace("__ATTACH_HINT__", hint)
            .replace("__SUBJECT__", html.escape(subject))
            .replace("__TO__", html.escape(", ".join(cfg.get("to", []))))
            .replace("__CC__", html.escape(", ".join(cfg.get("cc", []))))
            .replace("__BODY__", body_html))
    (out / "copy.html").write_text(page, encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("content")
    ap.add_argument("--config", default=str(HERE / "config.json"))
    ap.add_argument("--out", help="출력 폴더 (기본: content.json과 같은 폴더)")
    ap.add_argument("--attach", help="첨부할 대시보드 PDF 경로")
    a = ap.parse_args()

    cpath = Path(a.content).resolve()
    content = json.loads(cpath.read_text(encoding="utf-8"))
    cfg = json.loads(Path(a.config).read_text(encoding="utf-8"))
    out = Path(a.out).resolve() if a.out else cpath.parent
    out.mkdir(parents=True, exist_ok=True)
    attach = Path(a.attach).resolve() if a.attach else None
    if attach and not attach.exists():
        sys.exit(f"첨부 파일이 없습니다: {attach}")

    subject, body_html, body_txt, imgs, _ = build(content, cfg, cpath.parent, attach)

    # 미리보기/복사용 HTML
    page = (f'<!doctype html><html lang="ko"><head><meta charset="utf-8">'
            f"<title>{html.escape(subject)}</title></head>"
            f'<body style="background:#fff;margin:16px;">'
            f'<div style="{BASE}color:#666;border-bottom:1px solid #ddd;padding-bottom:8px;margin-bottom:12px;">'
            f"<b>제목</b> {html.escape(subject)}<br><b>받는 사람</b> {html.escape(', '.join(cfg.get('to', [])))}<br>"
            f"<b>참조</b> {html.escape(', '.join(cfg.get('cc', [])))}"
            + (f"<br><b>첨부</b> {html.escape(attach.name)}" if attach else "")
            + "<br><i>↓ 아래 본문만 복사해 메일 에디터에 붙여넣으세요.</i></div>"
            f"{body_html}</body></html>")
    (out / "mail.html").write_text(page, encoding="utf-8")
    (out / "mail.txt").write_text(f"제목: {subject}\n\n{body_txt}", encoding="utf-8")
    write_copy_page(out, subject, body_html, cfg, attach, content["month"])

    # EML (초안)
    msg = EmailMessage()
    msg["Subject"] = subject
    if cfg.get("from"):
        msg["From"] = cfg["from"]
    if cfg.get("to"):
        msg["To"] = ", ".join(cfg["to"])
    if cfg.get("cc"):
        msg["Cc"] = ", ".join(cfg["cc"])
    msg["X-Unsent"] = "1"  # Outlook에서 '보내기 전 초안'으로 열림
    msg.set_content(body_txt)
    eml_html = body_html
    for cid, path in imgs:
        eml_html = eml_html.replace(f'data-cid="{cid}" src="{path.as_uri()}"', f'src="cid:{cid}"')
    msg.add_alternative(eml_html, subtype="html")
    html_part = msg.get_payload()[1]
    for cid, path in imgs:
        maintype, subtype = (mimetypes.guess_type(path.name)[0] or "image/png").split("/")
        html_part.add_related(path.read_bytes(), maintype, subtype, cid=f"<{cid}>", filename=path.name)
    if attach:
        mt = (mimetypes.guess_type(attach.name)[0] or "application/octet-stream").split("/")
        msg.add_attachment(attach.read_bytes(), maintype=mt[0], subtype=mt[1], filename=attach.name)
    (out / "mail.eml").write_bytes(bytes(msg))

    print(f"제목: {subject}")
    for f in ("copy.html", "mail.html", "mail.eml", "mail.txt"):
        print(f"생성: {out / f}")


if __name__ == "__main__":
    main()
