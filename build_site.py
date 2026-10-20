# -*- coding: utf-8 -*-
"""把仓库内 PDF 报告提取重排为网页版，生成 GitHub Pages 站点。
结构：
  index.html            首页（报告卡片）
  site.css              全站样式
  reports/<slug>.html   每份报告的网页阅读版（末尾附下载链接）
  assets/<slug>/*.png   报告内图表
"""
import os, re, sys, html as htmllib
from urllib.parse import quote
sys.path.insert(0, os.path.expanduser("~/workspace/.pylibs"))
import pymupdf

BASE = os.path.dirname(os.path.abspath(__file__))

REPORTS = [
    dict(slug="2026-09-run", title="2026年9月中国经济运行报告",
         meta="2026年9月 · 经济运行 · 数据截至 2026-10-20", kind="pdf",
         pdf="9月/2026年9月中国经济运行报告.pdf",
         downloads=[("9月/2026年9月中国经济运行报告.pdf", "报告 PDF"),
                    ("9月/2026年9月经济运行数据表.xlsx", "数据表 XLSX")]),
    dict(slug="2026-08-fiscal", title="2026年8月中国财政与宏观经济数据报告",
         meta="2026年8月 · 财政与宏观 · 数据截至 2026-10-14", kind="pdf",
         pdf="8月/2026年8月中国财政与宏观经济数据报告.pdf",
         downloads=[("8月/2026年8月中国财政与宏观经济数据报告.pdf", "报告 PDF"),
                    ("8月/2026年8月财政与宏观经济数据表.xlsx", "数据表 XLSX")]),
    dict(slug="2026-08-run", title="2026年8月中国经济运行报告",
         meta="2026年8月 · 经济运行", kind="pdf",
         pdf="8月/2026年8月中国经济运行报告.pdf",
         downloads=[("8月/2026年8月中国经济运行报告.pdf", "报告 PDF"),
                    ("8月/2026年8月经济运行数据表.xlsx", "数据表 XLSX")]),
]

CSS = """
:root{color-scheme:light dark}
*{box-sizing:border-box}
body{margin:0;font-family:"Noto Sans CJK SC","PingFang SC","Microsoft YaHei",sans-serif;background:#f7f8fa;color:#1c1e21;line-height:1.85}
@media (prefers-color-scheme:dark){body{background:#14161a;color:#e8e8e8}}
.wrap{max-width:800px;margin:0 auto;padding:40px 20px 64px}
a{color:#0b3d91}
@media (prefers-color-scheme:dark){a{color:#7db3f5}}
.back{display:inline-block;margin-bottom:22px;font-size:14px;text-decoration:none}
.back:hover{text-decoration:underline}
header.doc{border-bottom:3px solid #0b3d91;padding-bottom:16px;margin-bottom:28px}
header.doc h1{font-size:27px;line-height:1.45;margin:0 0 8px}
header.doc .meta{color:#777;font-size:13.5px}
@media (prefers-color-scheme:dark){header.doc .meta{color:#9aa0a6}}
article h2{font-size:20px;margin:34px 0 12px;padding-bottom:6px;border-bottom:2px solid #0b3d91}
article h3{font-size:16.5px;margin:24px 0 10px;color:#0b3d91}
@media (prefers-color-scheme:dark){article h3{color:#7db3f5}}
article p{margin:10px 0;font-size:15.5px;text-align:justify}
article p.kpi{font-size:22px;font-weight:700;margin:14px 0 2px;color:#0b3d91}
@media (prefers-color-scheme:dark){article p.kpi{color:#7db3f5}}
article p.note,article .src{font-size:12.5px;color:#888;text-align:left}
article figure{margin:18px 0;text-align:center}
article figure img{max-width:100%;height:auto;border:1px solid #e3e6ea;border-radius:8px;background:#fff}
@media (prefers-color-scheme:dark){article figure img{border-color:#33363d}}
article table{border-collapse:collapse;width:100%;margin:14px 0;font-size:13px}
article th,article td{border:1px solid #c9ced6;padding:6px 9px;text-align:center}
article th{background:#eef2f7;font-weight:700}
@media (prefers-color-scheme:dark){article th{background:#232833}article th,article td{border-color:#3a3f4a}}
article td:first-child{text-align:left}
article ul,article ol{font-size:15.5px;margin:10px 0;padding-left:24px}
article li{margin-bottom:6px}
.dl{margin-top:44px;border-top:2px solid #0b3d91;padding-top:18px}
.dl h2{font-size:17px;margin:0 0 12px}
.dl a.btn{display:inline-block;margin:0 10px 10px 0;padding:9px 18px;border:1.5px solid #0b3d91;border-radius:8px;text-decoration:none;font-size:14.5px}
.dl a.btn:hover{background:#0b3d91;color:#fff}
footer.site{text-align:center;color:#999;font-size:13px;margin-top:40px}
/* 首页卡片 */
header.home h1{font-size:30px;margin:0 0 8px}
header.home p{color:#666;margin:0 0 34px;font-size:14.5px}
@media (prefers-color-scheme:dark){header.home p{color:#9aa0a6}}
.card{background:#fff;border:1px solid #e3e6ea;border-radius:12px;padding:20px 24px;margin-bottom:18px}
@media (prefers-color-scheme:dark){.card{background:#1d2026;border-color:#33363d}}
.card h2{font-size:19px;margin:0 0 6px}
.card h2 a{text-decoration:none;color:inherit}
.card h2 a:hover{color:#0b3d91;text-decoration:underline}
.card .meta{color:#888;font-size:13px;margin-bottom:10px}
.card .links a{font-size:14px;margin-right:16px;text-decoration:none}
.card .links a:hover{text-decoration:underline}
.card .links .read{font-weight:700}
"""

TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title} · 中国宏观经济报告</title>
<link rel="stylesheet" href="../site.css"></head>
<body><div class="wrap">
<a class="back" href="../index.html">← 返回报告列表</a>
<header class="doc"><h1>{title}</h1><div class="meta">{meta}</div></header>
<article>
{body}
</article>
<div class="dl"><h2>下载原文件</h2>{downloads}</div>
<footer class="site">中国宏观经济报告 · 网页版由 PDF 重排生成，数据以原文件为准</footer>
</div></body></html>
"""

INDEX_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>中国宏观经济报告</title>
<link rel="stylesheet" href="site.css"></head>
<body><div class="wrap">
<header class="home"><h1>中国宏观经济报告</h1>
<p>月度中国经济运行报告、财政与宏观经济数据报告及专题调研。点击「在线阅读」直接在网页查看全文，PDF 与数据表在每篇末尾下载。</p></header>
{cards}
<footer class="site">归档仓库：maijia765/china-macro-economy-reports</footer>
</div></body></html>
"""

def esc(t):
    return htmllib.escape(t, quote=False)

SPACED_RE = re.compile(r"^([A-Z] )+[A-Z]?\s*$")

def block_html(block):
    """把一个文本 block 转成 (分类, html)。分类: h2/h3/kpi/note/p/title/skip"""
    spans = []
    for line in block["lines"]:
        for s in line["spans"]:
            spans.append(s)
    if not spans:
        return "skip", ""
    text_all = "".join(s["text"] for s in spans).strip()
    if not text_all:
        return "skip", ""
    if SPACED_RE.match(text_all):
        return "skip", ""
    if re.fullmatch(r"[·•・.\s…]+", text_all):
        return "skip", ""
    maxsize = max(s["size"] for s in spans)
    # 页脚：纯页码
    if re.fullmatch(r"\d{1,3}", text_all) and maxsize <= 8.5:
        return "skip", ""
    # 独立序号标记（如结论区的 1. 2. 3.），并入下一段
    if re.fullmatch(r"\d+\.", text_all):
        return "marker", text_all
    parts = []
    for line in block["lines"]:
        for s in line["spans"]:
            t = esc(s["text"])
            if "Bold" in s["font"]:
                parts.append(f"<strong>{t}</strong>")
            else:
                parts.append(t)
    inner = "".join(parts).strip()
    if maxsize >= 24:
        return "title", inner
    if maxsize >= 17:
        return "kpi", f'<p class="kpi">{inner}</p>'
    if maxsize >= 14:
        return "h2", f"<h2>{inner}</h2>"
    if maxsize >= 10.8 and any("Bold" in s["font"] for s in spans):
        return "h3", f"<h3>{inner}</h3>"
    if maxsize <= 7.9:
        return "note", f'<p class="note">{inner}</p>'
    return "p", f"<p>{inner}</p>"

def page_spans(page):
    spans = []
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for s in line["spans"]:
                spans.append(s)
    return spans

def cell_text(cell, spans):
    """按坐标从原始 span 组装单元格文本（find_tables.extract 会把字符弄乱序）。"""
    if cell is None:
        return ""
    frags = []
    for s in spans:
        cx = (s["bbox"][0] + s["bbox"][2]) / 2
        cy = (s["bbox"][1] + s["bbox"][3]) / 2
        if cell[0] - 1 <= cx <= cell[2] + 1 and cell[1] - 1 <= cy <= cell[3] + 1:
            frags.append((round(cy / 5), s["bbox"][0], s["text"]))
    frags.sort()
    lines = {}
    for gy, _, t in frags:
        lines.setdefault(gy, []).append(t)
    return "".join("".join(lines[k]) for k in sorted(lines)).strip()

def table_html(tab, spans):
    """返回 (html, 网格是否可疑)。用 tab.rows 的按行单元格（tab.cells 不是完整网格）。"""
    if not tab.rows:
        return "", True
    grid = []
    for rowobj in tab.rows:
        grid.append([cell_text(c, spans) for c in rowobj.cells])
    # 逻辑行合并：检测网格会把跨行单元格按文本行拆成多个微行，
    # 首列为空的行是上一逻辑行的续行，按列拼接回去。
    if len(grid) > 1:
        merged = [grid[0]]
        for row in grid[1:]:
            if not any(row):
                continue
            if row[0] or len(merged) <= 1:
                merged.append(list(row))
            else:
                for i, t in enumerate(row):
                    if t:
                        merged[-1][i] = (merged[-1][i] + t) if merged[-1][i] else t
        grid = merged
    header = grid[0]
    nonempty = [i for i, t in enumerate(header) if t]
    suspicious = False
    if len(nonempty) >= 2:
        lo, hi = min(nonempty), max(nonempty)
        if any(not header[i] for i in range(lo, hi + 1)):
            suspicious = True
    out = ["<table>"]
    for r, row in enumerate(grid):
        if not any(row):
            continue
        tag = "th" if r == 0 else "td"
        out.append("<tr>" + "".join(f"<{tag}>{esc(t)}</{tag}>" for t in row) + "</tr>")
    out.append("</table>")
    return ("".join(out) if len(out) > 2 else ""), suspicious

def strip_tags(h):
    return re.sub(r"<[^>]+>", "", h)

def regroup_kpi(items):
    """财政版封面是 KPI 卡片网格（按行读会数字/标签分离），把卡片区按列重组。"""
    start = None
    for i, (_, _, h) in enumerate(items):
        if isinstance(h, str) and 'class="kpi"' in h:
            start = i
            break
    if start is None:
        return items
    end = start
    while end < len(items):
        h = items[end][2]
        if not isinstance(h, str):
            break
        if 'class="kpi"' in h or 'class="note"' in h or (h.startswith("<p>") and len(strip_tags(h)) <= 40):
            end += 1
        else:
            break
    window = items[start:end]
    if len(window) < 3:
        return items
    clusters = []
    for it in sorted(window, key=lambda x: x[1]):
        if clusters and it[1] - clusters[-1][-1][1] <= 25:
            clusters[-1].append(it)
        else:
            clusters.append([it])
    regrouped = []
    for cl in clusters:
        regrouped.extend(sorted(cl, key=lambda x: x[0]))
    return items[:start] + regrouped + items[end:]

def extract_pdf(pdf_path, slug):
    doc = pymupdf.open(pdf_path)
    asset_dir = os.path.join(BASE, "assets", slug)
    os.makedirs(asset_dir, exist_ok=True)
    body_parts = []
    img_n = 0
    tab_img_n = [0]
    seen_titles = 0
    for pno, page in enumerate(doc):
        items = []  # (y0, x0, html)
        spans = page_spans(page)
        tables = page.find_tables().tables
        table_bboxes = []
        for tab in tables:
            thtml, suspicious = table_html(tab, spans)
            if not thtml and not suspicious:
                continue
            table_bboxes.append(tab.bbox)
            if suspicious:
                # 网格对不齐：截取原 PDF 该区域为图片，保证数字零误差
                tab_img_n[0] += 1
                name = f"table-{tab_img_n[0]:02d}.png"
                clip = pymupdf.Rect(tab.bbox)
                pix = page.get_pixmap(matrix=pymupdf.Matrix(2.4, 2.4), clip=clip)
                pix.save(os.path.join(asset_dir, name))
                items.append((tab.bbox[1], tab.bbox[0],
                              f'<figure><img src="../assets/{slug}/{name}" alt="数据表" loading="lazy"></figure>'))
            else:
                items.append((tab.bbox[1], tab.bbox[0], thtml))
        def in_table(bbox):
            cx, cy = (bbox[0] + bbox[2]) / 2, (bbox[1] + bbox[3]) / 2
            return any(t[0] - 2 <= cx <= t[2] + 2 and t[1] - 2 <= cy <= t[3] + 2 for t in table_bboxes)
        def frac_in_table(block):
            ss = [s for l in block["lines"] for s in l["spans"]]
            if not ss:
                return 0.0
            inside = sum(1 for s in ss if in_table(s["bbox"]))
            return inside / len(ss)
        d = page.get_text("dict")
        for block in d["blocks"]:
            if block["type"] == 1:
                continue  # 图片走 get_image_info
            bbox = block["bbox"]
            if bbox[1] >= 795:  # 页脚区
                continue
            if frac_in_table(block) >= 0.5:
                continue
            kind, h = block_html(block)
            if kind == "skip" or not h:
                continue
            if kind == "title":
                seen_titles += 1
                continue  # 标题由模板呈现
            if kind == "marker":
                items.append((bbox[1], bbox[0], ("__MARKER__", h)))
                continue
            items.append((bbox[1], bbox[0], h))
        for info in page.get_image_info(xrefs=True):
            xref = info["xref"]
            bb = info["bbox"]
            if bb[1] >= 795 or in_table(bb):
                continue
            try:
                pix = pymupdf.Pixmap(doc, xref)
            except Exception:
                continue
            if pix.width < 200 or pix.height < 80:
                continue
            if pix.n > 3:
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
            img_n += 1
            name = f"img-{img_n:02d}.png"
            pix.save(os.path.join(asset_dir, name))
            items.append((bb[1], bb[0], f'<figure><img src="../assets/{slug}/{name}" alt="图表 {img_n}" loading="lazy"></figure>'))
        items.sort(key=lambda x: (round(x[0] / 6), x[1]))
        items = regroup_kpi(items)
        merged = []
        pending_marker = None
        for _, _, h in items:
            if isinstance(h, tuple):  # 序号标记
                pending_marker = h[1]
                continue
            if pending_marker and isinstance(h, str) and h.startswith("<p>"):
                h = "<p>" + pending_marker + " " + h[3:]
                pending_marker = None
            merged.append(h)
        body_parts.extend(merged)
    return "\n".join(body_parts), img_n

def build_europe(slug):
    src = os.path.expanduser("~/workspace/europe-macro-research/report_body.html")
    with open(src, encoding="utf-8") as f:
        frag = f.read()
    # 去掉原封面
    frag = re.sub(r'<div class="cover">.*?</div>', "", frag, flags=re.S, count=1)
    # 图表改成本地文件引用
    chart_src = os.path.expanduser("~/workspace/europe-macro-research/charts")
    asset_dir = os.path.join(BASE, "assets", slug)
    os.makedirs(asset_dir, exist_ok=True)
    import shutil
    def fig_sub(m):
        name = m.group(1)
        shutil.copy(os.path.join(chart_src, name), os.path.join(asset_dir, name))
        return f'<figure><img src="../assets/{slug}/{name}" alt="数据图" loading="lazy"></figure>'
    frag = re.sub(r"\{\{FIG:([a-z0-9_]+\.png)\}\}", fig_sub, frag)
    frag = frag.replace('<div class="fig">', "<figure>").replace("</div>", "</figure>", 0)
    # 标题降级：h1->h2, h2->h3
    frag = frag.replace("<h1>", "<h2>").replace("</h1>", "</h2>")
    frag = frag.replace("<h2>", "<h3>").replace("</h2>", "</h3>")
    return frag

def main():
    with open(os.path.join(BASE, "site.css"), "w", encoding="utf-8") as f:
        f.write(CSS)
    os.makedirs(os.path.join(BASE, "reports"), exist_ok=True)
    cards = []
    for rep in REPORTS:
        slug = rep["slug"]
        if rep["kind"] == "pdf":
            body, nimg = extract_pdf(os.path.join(BASE, rep["pdf"]), slug)
            print(slug, "images:", nimg, "body chars:", len(body))
        else:
            body = build_europe(slug)
            print(slug, "europe body chars:", len(body))
        dls = "".join(
            f'<a class="btn" href="../{quote(p)}" download>{esc(label)}</a>'
            for p, label in rep["downloads"])
        page = TEMPLATE.format(title=esc(rep["title"]), meta=esc(rep["meta"]),
                               body=body, downloads=dls)
        with open(os.path.join(BASE, "reports", slug + ".html"), "w", encoding="utf-8") as f:
            f.write(page)
        dl_links = "".join(
            f'<a href="{quote(p)}" download>{esc(label)}</a>' for p, label in rep["downloads"])
        cards.append(
            f'<div class="card"><h2><a href="reports/{slug}.html">{esc(rep["title"])}</a></h2>'
            f'<div class="meta">{esc(rep["meta"])}</div>'
            f'<div class="links"><a class="read" href="reports/{slug}.html">在线阅读 →</a>{dl_links}</div></div>')
    with open(os.path.join(BASE, "index.html"), "w", encoding="utf-8") as f:
        f.write(INDEX_TEMPLATE.format(cards="\n".join(cards)))
    print("site built")

if __name__ == "__main__":
    main()
