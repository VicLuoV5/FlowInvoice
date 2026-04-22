"""把 使用手册.md 转成单文件 HTML，方便用浏览器 Ctrl+P 另存为 PDF。

用法：
    python docs/build_manual_html.py

产物：docs/使用手册.html
然后用 Edge / Chrome 打开，Ctrl+P → 目标选「另存为 PDF」→ 保存。
"""
import os
import markdown

HERE = os.path.dirname(os.path.abspath(__file__))
MD = os.path.join(HERE, "使用手册.md")
HTML = os.path.join(HERE, "使用手册.html")

CSS = """
@page {
    size: A4;
    margin: 2cm 2cm 2cm 2cm;
}

* { box-sizing: border-box; }

body {
    font-family: "Microsoft YaHei", "微软雅黑", -apple-system, sans-serif;
    color: #3D3929;
    line-height: 1.75;
    font-size: 14px;
    max-width: 780px;
    margin: 0 auto;
    padding: 40px 32px;
    background: #FFFFFF;
}

h1 {
    color: #D97757;
    font-size: 28px;
    border-bottom: 3px solid #D97757;
    padding-bottom: 12px;
    margin-top: 0;
    page-break-before: auto;
}

h2 {
    color: #D97757;
    font-size: 20px;
    border-left: 4px solid #D97757;
    padding-left: 12px;
    margin-top: 36px;
    page-break-after: avoid;
}

h3 {
    color: #3D3929;
    font-size: 16px;
    margin-top: 24px;
    page-break-after: avoid;
}

p, li {
    color: #3D3929;
    line-height: 1.8;
}

a {
    color: #D97757;
    text-decoration: none;
    word-break: break-all;
}

a:hover { text-decoration: underline; }

strong { color: #3D3929; }

code {
    background: #F5F2EB;
    color: #B8431D;
    padding: 2px 6px;
    border-radius: 3px;
    font-family: "Consolas", "Monaco", monospace;
    font-size: 0.9em;
}

pre {
    background: #2A271D;
    color: #FBEAE0;
    padding: 16px;
    border-radius: 6px;
    overflow-x: auto;
    line-height: 1.5;
    page-break-inside: avoid;
}

pre code {
    background: transparent;
    color: inherit;
    padding: 0;
}

blockquote {
    border-left: 4px solid #D97757;
    background: #FBEAE0;
    margin: 16px 0;
    padding: 12px 16px;
    color: #3D3929;
    border-radius: 0 4px 4px 0;
}

blockquote p { margin: 4px 0; }

img {
    max-width: 100%;
    border: 1px solid #E8E4DB;
    border-radius: 4px;
    margin: 12px 0;
    display: block;
    page-break-inside: avoid;
}

table {
    border-collapse: collapse;
    width: 100%;
    margin: 16px 0;
    page-break-inside: avoid;
}

th {
    background: #D97757;
    color: #FFFFFF;
    padding: 10px 12px;
    text-align: left;
    font-weight: bold;
}

td {
    border: 1px solid #E8E4DB;
    padding: 10px 12px;
}

tr:nth-child(even) td { background: #FBF8F2; }

hr {
    border: none;
    border-top: 1px solid #E8E4DB;
    margin: 32px 0;
}

ul, ol { padding-left: 24px; }
li { margin: 4px 0; }

/* 目录链接样式更紧凑 */
ul li a { color: #3D3929; }
ul li a:hover { color: #D97757; }

/* 打印优化 */
@media print {
    body {
        padding: 0;
        max-width: 100%;
    }
    a { color: #D97757; }
    pre, blockquote, img, table { page-break-inside: avoid; }
    h1, h2, h3 { page-break-after: avoid; }
}
"""

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>极简票流 FlowInvoice 使用手册</title>
<style>{css}</style>
</head>
<body>
{body}
<hr>
<p style="text-align:center; color:#8C8070; font-size:0.85em;">
极简票流 FlowInvoice · 纯本地离线 · MIT License<br>
<a href="https://github.com/VicLuoV5/FlowInvoice">https://github.com/VicLuoV5/FlowInvoice</a>
</p>
</body>
</html>
"""


def main():
    with open(MD, "r", encoding="utf-8") as f:
        md_text = f.read()

    body = markdown.markdown(
        md_text,
        extensions=["tables", "fenced_code", "toc", "nl2br"],
        output_format="html5",
    )

    html = HTML_TEMPLATE.format(css=CSS, body=body)

    with open(HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[OK] generated: {HTML}")
    print("next: open HTML in Edge/Chrome, Ctrl+P, target = Save as PDF")


if __name__ == "__main__":
    main()
