# app.py
# Streamlit: 제안 옵션 선택 → (요청별 인라인 미리보기) → PDF/Excel/PPT 내보내기
# - "요청별 옵션 선택" 라디오 바로 아래에 미리보기 카드 표시
# - 미리보기는 "제안요청 제목" 아래 줄에 "옵션 N — 대제목"을 가볍게 표시(하위 레벨 톤)
# - PDF/PPT/Excel 및 한글 폰트 자동탐지, 옵션 대제목 생성 그대로 유지

import os
import io
import json
import platform
from datetime import datetime
from typing import Optional, Tuple, List, Dict, Any

import pandas as pd
import numpy as np
import streamlit as st

# ====== PDF (ReportLab) ======
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
    Flowable
)
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ====== PPT (python-pptx) ======
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN

# =====================================================================================
# 유틸: 안전 문자열/리스트/JSON
# =====================================================================================

def S(x: Any) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and np.isnan(x):
        return ""
    return str(x)

def parse_url_list(val: Any) -> List[str]:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return []
    if isinstance(val, list):
        return [S(u).strip() for u in val if S(u).strip()]
    s = S(val)
    parts = []
    for tok in s.replace(";", "\n").split("\n"):
        for sub in tok.split(","):
            u = sub.strip()
            if u:
                parts.append(u)
    return parts

def parse_timeline(val: Any) -> List[Dict[str, Any]]:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return []
    if isinstance(val, list):
        return val
    s = S(val).strip()
    if not s:
        return []
    try:
        obj = json.loads(s)
        if isinstance(obj, list):
            return obj
    except Exception:
        pass
    return []

def try_extract_overview_table_from_row(row: pd.Series) -> Optional[Dict[str, Any]]:
    keys = ("columns", "rows")
    for col in row.index:
        v = row[col]
        if isinstance(v, dict) and all(k in v for k in keys):
            return v
        s = S(v).strip()
        if s.startswith("{") and s.endswith("}"):
            try:
                obj = json.loads(s)
                if isinstance(obj, dict) and all(k in obj for k in keys):
                    return obj
            except Exception:
                continue
    return None

# =====================================================================================
# 한글 폰트 자동 탐지/등록
# =====================================================================================

def _candidate_font_paths() -> list[Tuple[str, Optional[int], str]]:
    sys = platform.system()
    cands: list[Tuple[str, Optional[int], str]] = []

    env_path = os.getenv("KOREAN_TTF_PATH")
    if env_path and os.path.exists(env_path):
        idx = None
        if env_path.lower().endswith(".ttc"):
            try:
                idx = int(os.getenv("KOREAN_TTC_INDEX", "0"))
            except:
                idx = 0
        cands.append((env_path, idx, "KR-Body"))

    if sys == "Windows":
        win_fonts = r"C:\Windows\Fonts"
        cands += [
            (os.path.join(win_fonts, "malgun.ttf"), None, "MalgunGothic-Regular"),
            (os.path.join(win_fonts, "malgunbd.ttf"), None, "MalgunGothic-Bold"),
            (os.path.join(win_fonts, "NanumGothic.ttf"), None, "NanumGothic"),
            (os.path.join(win_fonts, "NotoSansKR-Regular.otf"), None, "NotoSansKR-Regular"),
        ]
    elif sys == "Darwin":
        cands += [
            ("/Library/Fonts/AppleSDGothicNeo.ttc", 0, "AppleSDGothicNeo-0"),
            ("/System/Library/Fonts/AppleSDGothicNeo.ttc", 0, "AppleSDGothicNeo-0"),
            ("/Library/Fonts/NanumGothic.ttf", None, "NanumGothic"),
            ("/Library/Fonts/NotoSansKR-Regular.otf", None, "NotoSansKR-Regular"),
        ]
    else:
        cands += [
            ("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", None, "NanumGothic"),
            ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0, "NotoSansCJK-Regular"),
            ("/usr/share/fonts/opentype/noto/NotoSansKR-Regular.otf", None, "NotoSansKR-Regular"),
            ("/usr/share/fonts/truetype/noto/NotoSansKR-Regular.otf", None, "NotoSansKR-Regular"),
        ]
    out, seen = [], set()
    for p, idx, name in cands:
        if p and os.path.exists(p):
            key = (p, idx, name)
            if key not in seen:
                seen.add(key)
                out.append(key)
    return out

def register_korean_font_for_pdf() -> Optional[str]:
    for path, idx, name in _candidate_font_paths():
        try:
            if path.lower().endswith(".ttc"):
                TT = TTFont(name, path, subfontIndex=(0 if idx is None else idx))
            else:
                TT = TTFont(name, path)
            pdfmetrics.registerFont(TT)
            return name
        except Exception as e:
            print(f"[PDF] Font register failed: {path} -> {e}")
            continue
    return None

def _ppt_ko_font_name() -> str:
    sys = platform.system()
    return {
        "Windows": "Malgun Gothic",
        "Darwin": "Apple SD Gothic Neo",
        "Linux": "Noto Sans CJK KR"
    }.get(sys, "Malgun Gothic")

PP_KO_FONT = _ppt_ko_font_name()
PDF_KO_FONT = register_korean_font_for_pdf()

# =====================================================================================
# PDF 스타일 + HR(구분선)
# =====================================================================================

def build_pdf_styles() -> Dict[str, ParagraphStyle]:
    styles = getSampleStyleSheet()
    base = PDF_KO_FONT or styles["Normal"].fontName
    if "K-Body" not in styles:
        styles.add(ParagraphStyle(
            name="K-Body", parent=styles["Normal"],
            fontName=base, fontSize=11, leading=14,
            spaceBefore=3, spaceAfter=4, textColor=colors.black
        ))
    if "K-H3" not in styles:
        styles.add(ParagraphStyle(
            name="K-H3", parent=styles["Normal"],
            fontName=base, fontSize=13, leading=16,
            spaceBefore=6, spaceAfter=3, textColor=colors.HexColor("#222")
        ))
    if "K-H2" not in styles:
        styles.add(ParagraphStyle(
            name="K-H2", parent=styles["Normal"],
            fontName=base, fontSize=15, leading=18,
            spaceBefore=8, spaceAfter=4, textColor=colors.HexColor("#111")
        ))
    if "K-H1" not in styles:
        styles.add(ParagraphStyle(
            name="K-H1", parent=styles["Normal"],
            fontName=base, fontSize=18, leading=22,
            spaceBefore=10, spaceAfter=6, textColor=colors.HexColor("#0D0D0D")
        ))
    if "K-Title" not in styles:
        styles.add(ParagraphStyle(
            name="K-Title", parent=styles["Title"],
            fontName=base, fontSize=22, leading=26,
            spaceBefore=8, spaceAfter=8, alignment=1, textColor=colors.HexColor("#0A0A0A")
        ))
    if "K-Label" not in styles:
        styles.add(ParagraphStyle(
            name="K-Label", parent=styles["Normal"],
            fontName=base, fontSize=9, leading=11, textColor=colors.HexColor("#666")
        ))
    return styles

class HR(Flowable):
    def __init__(self, width=1, thickness=0.5, color=colors.HexColor("#DDDDDD"), spaceBefore=6, spaceAfter=6):
        Flowable.__init__(self)
        self.width = width
        self.thickness = thickness
        self.color = color
        self.spaceBefore = spaceBefore
        self.spaceAfter = spaceAfter

    def wrap(self, availWidth, availHeight):
        self._w = availWidth if self.width == 1 else min(self.width, availWidth)
        return self._w, self.thickness + self.spaceBefore + self.spaceAfter

    def draw(self):
        self.canv.saveState()
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 0, self._w, 0)
        self.canv.restoreState()

PDF_STYLES = build_pdf_styles()

# =====================================================================================
# PPT 도우미
# =====================================================================================

def apply_ppt_text_style(shape, size_pt: int = 16, bold: bool = False, align: str = "left", line_spacing: float = 1.2):
    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        return
    tf = shape.text_frame
    if tf.paragraphs:
        if align == "center":
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        elif align == "right":
            tf.paragraphs[0].alignment = PP_ALIGN.RIGHT
        else:
            tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    for p in tf.paragraphs:
        try:
            p.line_spacing = line_spacing
        except:
            pass
        for r in p.runs:
            r.font.name = PP_KO_FONT
            r.font.size = Pt(size_pt)
            r.font.bold = bool(bold)

def add_textbox(slide, left_in, top_in, width_in, height_in, text="", size=16, bold=False, align="left", line_spacing=1.2):
    tx = slide.shapes.add_textbox(Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in))
    tf = tx.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = S(text)
    apply_ppt_text_style(tx, size_pt=size, bold=bold, align=align, line_spacing=line_spacing)
    return tx

def add_title_subtitle(slide, title, subtitle):
    title_box = add_textbox(slide, 0.8, 0.7, 11.0, 1.2, S(title), size=34, bold=True, align="left", line_spacing=1.1)
    if S(subtitle):
        subtitle_box = add_textbox(slide, 0.8, 1.7, 11.0, 0.8, S(subtitle), size=18, bold=False, align="left")
    else:
        subtitle_box = None
    return title_box, subtitle_box

def bullets_from_paragraphs(slide, left, top, width, height, lines: List[str], size=16):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.clear()
    first = True
    for line in lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = S(line)
        p.level = 0
    apply_ppt_text_style(tb, size_pt=size, bold=False, align="left", line_spacing=1.25)
    return tb

# =====================================================================================
# 옵션 대제목 생성
# =====================================================================================

def compute_option_big_titles(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "옵션대제목" not in df.columns:
        df["옵션대제목"] = ""
    for (rid, opt), g in df.groupby(["요청 ID", "옵션번호"]):
        if not S(opt).isdigit():
            continue
        existing = S(g["옵션대제목"].dropna().astype(str).head(1).tolist()[0] if not g["옵션대제목"].empty else "")
        if existing:
            big = existing
        else:
            big = ""
            meta = g[g["슬라이드번호"] == "META"]
            if not meta.empty:
                mt = S(meta.iloc[0].get("제목"))
                if mt:
                    big = mt
            if not big:
                detail = g[g["슬라이드번호"].apply(lambda v: S(v).isdigit())].copy()
                if not detail.empty:
                    detail["슬라이드번호"] = detail["슬라이드번호"].astype(int)
                    detail = detail.sort_values("슬라이드번호")
                    big = S(detail.iloc[0].get("제목"))
            if not big:
                big = f"옵션 {opt}"
        df.loc[(df["요청 ID"] == rid) & (df["옵션번호"] == opt), "옵션대제목"] = big
    return df

# =====================================================================================
# PDF 생성
# =====================================================================================

def build_pdf(selected_df: pd.DataFrame, client_info: Dict[str, str], body_size=11) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm, topMargin=14*mm, bottomMargin=14*mm
    )
    styles = PDF_STYLES
    styles["K-Body"].fontSize = body_size
    styles["K-Body"].leading = int(body_size * 1.3)

    story: List[Any] = []

    title = f"{S(client_info.get('고객사',''))} 제안 옵션 패키지"
    sub = f"{S(client_info.get('작성팀',''))} · {S(client_info.get('작성일',''))}"
    story += [
        Spacer(1, 18),
        Paragraph(title, styles["K-Title"]),
        Paragraph(sub, styles["K-Label"]),
        HR(),
    ]

    summary_rows = []
    for req_id, grp in selected_df.groupby("요청 ID"):
        if req_id in ("COVER", "CLOSING"):
            continue
        req_title = S(grp["요청 제목"].iloc[0] if "요청 제목" in grp.columns else req_id)
        opt = ""
        big = ""
        opts = [x for x in grp["옵션번호"].unique().tolist() if S(x).isdigit()]
        if opts:
            opt = S(opts[0])
            g2 = grp[grp["옵션번호"] == opt]
            if not g2.empty and "옵션대제목" in g2.columns:
                big = S(g2["옵션대제목"].iloc[0])
        summary_rows.append([S(req_id), req_title, opt, big])
    if summary_rows:
        data = [["요청 ID", "요청 제목", "선택 옵션", "옵션 대제목"]] + summary_rows
        colw = [22*mm, 98*mm, 18*mm, 44*mm]
        tbl = Table(data, hAlign='LEFT', colWidths=colw)
        tbl.setStyle(TableStyle([
            ('FONTNAME', (0,0), (-1,-1), PDF_KO_FONT or 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,0), 11),
            ('FONTSIZE', (0,1), (-1,-1), 9.8),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#F6F6F6")),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('ALIGN', (0,1), (0,-1), 'CENTER'),
            ('ALIGN', (2,1), (2,-1), 'CENTER'),
            ('INNERGRID', (0,0), (-1,-1), 0.25, colors.HexColor("#DDD")),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CCC")),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story += [tbl, HR()]

    ordered_ids = [x for x in selected_df["요청 ID"].unique() if x not in ("COVER","CLOSING")]

    for idx_req, req_id in enumerate(ordered_ids):
        grp = selected_df[selected_df["요청 ID"] == req_id]
        req_title = S(grp["요청 제목"].iloc[0] if "요청 제목" in grp.columns else req_id)

        story += [
            Paragraph(f"[{S(req_id)}] {req_title}", styles["K-H1"]),
            HR()
        ]

        sel_opts = [x for x in grp["옵션번호"].unique().tolist() if S(x).isdigit()]
        sel = S(sel_opts[0]) if sel_opts else ""
        big = ""
        if sel:
            gg = grp[grp["옵션번호"] == sel]
            if not gg.empty and "옵션대제목" in gg.columns:
                big = S(gg["옵션대제목"].iloc[0])

        if sel:
            story.append(Paragraph(f"옵션 {sel} · {big}", styles["K-H2"]))
            story.append(HR(color=colors.HexColor("#EEEEEE")))

        over = grp[grp["슬라이드번호"] == "OVERVIEW"]
        if not over.empty:
            ov = over.iloc[0]
            ov_tab = try_extract_overview_table_from_row(ov)
            if ov_tab:
                cols = ov_tab.get("columns", [])
                rows = ov_tab.get("rows", [])
                data = [cols] + rows if cols and rows else []
                if data:
                    t = Table(data, hAlign='LEFT')
                    t.setStyle(TableStyle([
                        ('FONTNAME', (0,0), (-1,-1), PDF_KO_FONT or 'Helvetica'),
                        ('FONTSIZE', (0,0), (-1,0), 10.5),
                        ('FONTSIZE', (0,1), (-1,-1), 9.5),
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#F2F2F2")),
                        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.HexColor("#E1E1E1")),
                        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#D0D0D0")),
                        ('TOPPADDING', (0,0), (-1,-1), 3),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
                    ]))
                    story += [t, HR()]

        meta = grp[grp["슬라이드번호"] == "META"]
        if not meta.empty:
            m = meta.iloc[0]
            parts = [
                ("왜 이 옵션인가", m.get("왜_이_옵션")),
                ("적합 시그널", m.get("적합_시그널")),
                ("리스크", m.get("리스크")),
                ("완화책", m.get("완화책")),
            ]
            for i, (h, b) in enumerate(parts):
                story.append(Paragraph(S(h), styles["K-H3"]))
                if S(b):
                    for ln in [x.strip() for x in S(b).split("\n") if x.strip()]:
                        story.append(Paragraph(ln, styles["K-Body"]))
                if i < len(parts) - 1:
                    story.append(HR())

            tl = parse_timeline(m.get("타임라인"))
            if tl:
                story += [Paragraph("타임라인(주)", styles["K-H3"])]
                tidata = [["Phase", "기간(주)"]] + [[S(x.get("phase")), S(x.get("duration_weeks"))] for x in tl]
                from reportlab.platypus import Table
                tt = Table(tidata, hAlign='LEFT', colWidths=[110*mm, 30*mm])
                tt.setStyle(TableStyle([
                    ('FONTNAME', (0,0), (-1,-1), PDF_KO_FONT or 'Helvetica'),
                    ('FONTSIZE', (0,0), (-1,0), 10.5),
                    ('FONTSIZE', (0,1), (-1,-1), 9.5),
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#F7F7F7")),
                    ('INNERGRID', (0,0), (-1,-1), 0.25, colors.HexColor("#DDDDDD")),
                    ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CCCCCC")),
                ]))
                story += [tt, HR(color=colors.HexColor("#EEEEEE"))]

        detail = grp[grp["슬라이드번호"].apply(lambda v: S(v).isdigit())].copy()
        if not detail.empty:
            detail["슬라이드번호"] = detail["슬라이드번호"].astype(int)
            detail = detail.sort_values(by=["슬라이드번호"])
            for j, (_, r) in enumerate(detail.iterrows()):
                story.append(Paragraph(S(r.get("제목")), styles["K-H2"]))
                if S(r.get("부제목")):
                    story.append(Paragraph(S(r.get("부제목")), styles["K-H3"]))
                body = S(r.get("본문초안"))
                if body:
                    for ln in [x.strip() for x in body.split("\n") if x.strip()]:
                        story.append(Paragraph(ln, styles["K-Body"]))

                urls = parse_url_list(r.get("URL"))
                if urls:
                    story.append(Paragraph("참고 URL", styles["K-H3"]))
                    for u in urls:
                        story.append(Paragraph(S(u), styles["K-Body"]))

                if j < len(detail) - 1:
                    story.append(HR())

        if idx_req < len(ordered_ids) - 1:
            story.append(PageBreak())

    doc.build(story)
    return buf.getvalue()

# =====================================================================================
# PPT 생성
# =====================================================================================

def build_ppt(selected_df: pd.DataFrame, client_info: Dict[str, str], body_size=16) -> bytes:
    prs = Presentation()
    blank = prs.slide_layouts[6]

    cover = prs.slides.add_slide(blank)
    title = f"{S(client_info.get('고객사',''))} 제안 옵션 패키지"
    subtitle = f"{S(client_info.get('작성팀',''))} · {S(client_info.get('작성일',''))}"
    add_title_subtitle(cover, title, subtitle)

    for req_id, grp in selected_df.groupby("요청 ID"):
        if req_id in ("COVER", "CLOSING"):
            continue
        req_title = S(grp["요청 제목"].iloc[0] if "요청 제목" in grp.columns else req_id)

        sel_opts = [x for x in grp["옵션번호"].unique().tolist() if S(x).isdigit()]
        sel = S(sel_opts[0]) if sel_opts else ""
        big = ""
        if sel:
            gg = grp[grp["옵션번호"] == sel]
            if not gg.empty and "옵션대제목" in gg.columns:
                big = S(gg["옵션대제목"].iloc[0])

        s = prs.slides.add_slide(blank)
        add_title_subtitle(s, f"[{S(req_id)}] {req_title}", f"옵션 {sel} · {big}")

        meta = grp[grp["슬라이드번호"] == "META"]
        if not meta.empty:
            m = meta.iloc[0]
            y = 2.2
            add_textbox(s, 0.9, y, 10.6, 0.5, f"옵션 {sel} · {big}", size=22, bold=True)
            y += 0.7
            bullets = []
            if S(m.get("왜_이_옵션")):
                bullets.append("• " + S(m.get("왜_이_옵션")).split("\n")[0])
            if S(m.get("적합_시그널")):
                bullets.append("• " + S(m.get("적합_시그널")).split("\n")[0])
            if S(m.get("리스크")):
                bullets.append("• " + S(m.get("리스크")).split("\n")[0])
            if S(m.get("완화책")):
                bullets.append("• " + S(m.get("완화책")).split("\n")[0])
            if not bullets:
                bullets = ["• 요약 정보"]
            bullets_from_paragraphs(s, 0.9, y, 10.6, 3.5, bullets, size=body_size)

        detail = grp[grp["슬라이드번호"].apply(lambda v: S(v).isdigit())].copy()
        if not detail.empty:
            detail["슬라이드번호"] = detail["슬라이드번호"].astype(int)
            detail = detail.sort_values(by=["슬라이드번호"])
            for _, r in detail.iterrows():
                ss = prs.slides.add_slide(blank)
                add_title_subtitle(ss, S(r.get("제목")), f"옵션 {sel} · {big}")
                body = S(r.get("본문초안"))
                if body:
                    lines = [ln.strip() for ln in body.split("\n") if ln.strip()]
                    bullets_from_paragraphs(ss, 0.9, 2.2, 10.6, 4.8, lines, size=body_size)
                urls = parse_url_list(r.get("URL"))
                if urls:
                    add_textbox(ss, 0.9, 7.3, 10.6, 0.5, "참고 URL", size=14, bold=True)
                    add_textbox(ss, 0.9, 7.8, 10.6, 0.8, "\n".join(urls), size=12)

    closing = prs.slides.add_slide(blank)
    add_title_subtitle(closing, "다음 단계", "")
    bullets_from_paragraphs(closing, 0.9, 2.2, 10.6, 3.0, [
        "옵션 선택 워크숍",
        "데이터/사전조건 점검",
        "파일럿 범위 합의 및 킥오프"
    ], size=body_size)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()

# =====================================================================================
# Excel 생성
# =====================================================================================

def build_excel(selected_df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        selected_df.to_excel(writer, index=False, sheet_name="SelectedOptions")
        wb = writer.book
        ws = writer.sheets["SelectedOptions"]
        fmt = wb.add_format({"font_name": "Malgun Gothic", "font_size": 10})
        ws.set_column(0, selected_df.shape[1]-1, 24, fmt)
    return out.getvalue()

# =====================================================================================
# 인라인 미리보기 렌더러 (요청별 카드)
# =====================================================================================

def render_inline_preview(req_id: str, sub_df: pd.DataFrame, selected_opt: str):
    sub_df = sub_df.copy()
    req_title = S(sub_df["요청 제목"].iloc[0] if "요청 제목" in sub_df.columns and not sub_df.empty else req_id)
    opt_df = sub_df[sub_df["옵션번호"] == selected_opt]
    big = S(opt_df["옵션대제목"].iloc[0]) if not opt_df.empty and "옵션대제목" in opt_df.columns else ""

    # 상단: 제안요청 제목(굵게), 아래 줄: 옵션 N — 대제목(캡션 톤)
    st.markdown(f"**{req_title}**")
    st.caption(f"선택: 옵션 {selected_opt} — {big}")

    # OVERVIEW 비교표는 확장형으로
    ov = sub_df[sub_df["슬라이드번호"] == "OVERVIEW"]
    if not ov.empty:
        ov_tab = try_extract_overview_table_from_row(ov.iloc[0])
        if ov_tab and ov_tab.get("columns") and ov_tab.get("rows"):
            with st.expander("옵션 비교(OVERVIEW)", expanded=False):
                st.table(pd.DataFrame(ov_tab["rows"], columns=ov_tab["columns"]))

    # META 요약(각 항목 첫 줄만 간략 표기)
    meta = opt_df[opt_df["슬라이드번호"] == "META"]
    if not meta.empty:
        m = meta.iloc[0]
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**왜 이 옵션인가**")
            msg = S(m.get("왜_이_옵션"))
            if msg:
                st.write(msg.split("\n")[0])
            st.markdown("**적합 시그널**")
            msg = S(m.get("적합_시그널"))
            if msg:
                st.write(msg.split("\n")[0])
        with col2:
            st.markdown("**리스크**")
            msg = S(m.get("리스크"))
            if msg:
                st.write(msg.split("\n")[0])
            st.markdown("**완화책**")
            msg = S(m.get("완화책"))
            if msg:
                st.write(msg.split("\n")[0])

        tl = parse_timeline(m.get("타임라인"))
        if tl:
            with st.expander("타임라인(주)", expanded=False):
                st.table(pd.DataFrame(tl))

    # 상세: 첫 1~2개만 미리보기(짧게), 전체는 확장
    detail = opt_df[opt_df["슬라이드번호"].apply(lambda v: S(v).isdigit())].copy()
    if not detail.empty:
        detail["슬라이드번호"] = detail["슬라이드번호"].astype(int)
        detail = detail.sort_values("슬라이드번호")
        top_n = detail.head(2)
        for _, r in top_n.iterrows():
            st.markdown(f"- **{S(r.get('제목'))}**")
            b = S(r.get("본문초안"))
            if b:
                st.write("  " + b.split("\n")[0])

        with st.expander("상세 슬라이드 전체 보기", expanded=False):
            for _, r in detail.iterrows():
                st.markdown(f"**{S(r.get('제목'))}**")
                if S(r.get("부제목")):
                    st.caption(S(r.get("부제목")))
                body = S(r.get("본문초안"))
                if body:
                    for ln in [x.strip() for x in body.split("\n") if x.strip()]:
                        st.write("- " + ln)
                urls = parse_url_list(r.get("URL"))
                if urls:
                    st.caption("참고 URL")
                    for u in urls:
                        st.write(u)
                st.divider()

# =====================================================================================
# Streamlit UI
# =====================================================================================

st.set_page_config(page_title="Proposal Options Builder", layout="wide")
st.title("제안 옵션 선택 · 미리보기 · 내보내기")

with st.sidebar:
    st.subheader("내보내기")
    pdf_body_size = st.slider("PDF 본문 글자 크기", 9, 14, 11)
    ppt_body_size = st.slider("PPT 본문 글자 크기", 12, 20, 16)
    st.markdown("---")
    st.caption(f"OS: {platform.system()} | PDF Font: {PDF_KO_FONT or '기본'} | PPT Font: {_ppt_ko_font_name()}")

st.markdown("#### 1) 데이터 업로드")
uploaded = st.file_uploader("`slim_master_slide` CSV/Excel 업로드", type=["csv", "xlsx"])
if uploaded is None:
    st.info("CSV/Excel를 업로드하세요. (필수 컬럼 예시: 요청 ID, 요청 제목, 옵션번호, 슬라이드번호, 제목, 부제목, 본문초안, 왜_이_옵션, 적합_시그널, 리스크, 완화책, 타임라인, URL, (선택) 옵션대제목)")
    st.stop()

if uploaded.name.lower().endswith(".csv"):
    df = pd.read_csv(uploaded, dtype=str).fillna("")
else:
    df = pd.read_excel(uploaded, dtype=str).fillna("")

required_cols = ["요청 ID","요청 제목","옵션번호","슬라이드번호","제목","부제목","본문초안","왜_이_옵션","적합_시그널","리스크","완화책","타임라인","URL"]
for c in required_cols:
    if c not in df.columns:
        df[c] = ""

df = compute_option_big_titles(df)
df["슬라이드번호"] = df["슬라이드번호"].astype(str)
df["옵션번호"] = df["옵션번호"].astype(str)

st.markdown("#### 2) 고객 정보")
col_a, col_b, col_c = st.columns(3)
with col_a:
    client_name = st.text_input("고객사", value="")
with col_b:
    author = st.text_input("작성팀", value="")
with col_c:
    today_str = datetime.now().strftime("%Y-%m-%d")
    date_str = st.text_input("작성일", value=today_str)

client_info = {"고객사": client_name, "작성팀": author, "작성일": date_str}

st.markdown("#### 3) 요청별 옵션 선택 (아래에 즉시 미리보기)")
req_ids = [x for x in df["요청 ID"].unique().tolist() if x not in ("COVER","CLOSING")]
sel_map: Dict[str, str] = {}

for rid in req_ids:
    sub = df[df["요청 ID"] == rid]
    req_title = S(sub["요청 제목"].iloc[0] if not sub.empty else rid)
    opts = sorted({o for o in sub["옵션번호"].unique().tolist() if S(o).isdigit()}, key=lambda x: int(x) if S(x).isdigit() else 999)
    if not opts:
        continue

    # 라디오에 "옵션 N — 대제목" 형태
    big_title_map = {}
    for o in opts:
        g = sub[sub["옵션번호"] == o]
        bt = S(g["옵션대제목"].iloc[0]) if not g.empty and "옵션대제목" in g.columns else ""
        big_title_map[o] = bt

    st.markdown(f"**[{rid}] {req_title}**")
    sel = st.radio(
        "옵션을 선택하세요",
        options=opts,
        horizontal=True,
        index=0,
        key=f"sel_{rid}",
        format_func=lambda o: f"{o} — {big_title_map.get(o, '')}" if big_title_map.get(o, "") else f"{o}"
    )
    sel_map[rid] = sel

    # 👉 선택 직후 인라인 미리보기 카드
    with st.container():
        render_inline_preview(rid, sub, sel)
    st.divider()

# 선택 데이터 구성(내보내기용)
frames = []
cover_rows = df[df["요청 ID"] == "COVER"]
if not cover_rows.empty:
    frames.append(cover_rows)
for rid, sel in sel_map.items():
    sub = df[df["요청 ID"] == rid]
    overview = sub[sub["슬라이드번호"] == "OVERVIEW"]
    if not overview.empty:
        frames.append(overview)
    part = sub[sub["옵션번호"] == sel]
    frames.append(part)
closing_rows = df[df["요청 ID"] == "CLOSING"]
if not closing_rows.empty:
    frames.append(closing_rows)
selected_df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=df.columns)

st.markdown("#### 4) 내보내기")
col1, col2, col3 = st.columns(3)
with col1:
    if st.button("📄 PDF 생성", use_container_width=True):
        try:
            pdf_bytes = build_pdf(selected_df, client_info, body_size=pdf_body_size)
            st.success("PDF 생성 완료")
            st.download_button("PDF 다운로드", data=pdf_bytes, file_name="proposal_options.pdf",
                               mime="application/pdf", use_container_width=True)
        except Exception as e:
            st.error(f"PDF 생성 오류: {e}")

with col2:
    if st.button("📊 Excel 생성", use_container_width=True):
        try:
            xlsx_bytes = build_excel(selected_df)
            st.success("Excel 생성 완료")
            st.download_button("Excel 다운로드", data=xlsx_bytes, file_name="proposal_options.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        except Exception as e:
            st.error(f"Excel 생성 오류: {e}")

with col3:
    if st.button("🖼️ PPT 생성", use_container_width=True):
        try:
            ppt_bytes = build_ppt(selected_df, client_info, body_size=ppt_body_size)
            st.success("PPT 생성 완료")
            st.download_button("PPT 다운로드", data=ppt_bytes, file_name="proposal_options.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               use_container_width=True)
        except Exception as e:
            st.error(f"PPT 생성 오류: {e}")

with st.expander("선택 데이터 미리보기(전체)", expanded=False):
    st.dataframe(selected_df, use_container_width=True)
