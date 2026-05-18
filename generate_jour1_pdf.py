import markdown
from weasyprint import HTML
from pathlib import Path

md_content = Path("/home/user/fatou/JOUR1_Lundi_Bienvenue_IA.md").read_text(encoding="utf-8")

html_body = markdown.markdown(md_content, extensions=["tables", "fenced_code"])

html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<style>
  body {{
    font-family: Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.75;
    color: #1a1a2e;
    margin: 0;
    padding: 0;
  }}

  @page {{
    margin: 2.2cm 2cm 2cm 2cm;
    @bottom-right {{
      content: "TAMOU NEURAL PATH  |  Jour 1  |  Page " counter(page) " / " counter(pages);
      font-size: 8.5pt;
      color: #999;
    }}
    @top-right {{
      content: "Lundi — Bienvenue dans l'IA";
      font-size: 8.5pt;
      color: #999;
    }}
  }}

  /* TITRE PRINCIPAL */
  h1:first-of-type {{
    font-size: 28pt;
    font-weight: 900;
    color: #ffffff;
    background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
    text-align: center;
    padding: 45px 30px 35px 30px;
    margin: 0 0 0 0;
    letter-spacing: 3px;
    page-break-after: avoid;
  }}

  h1 {{
    font-size: 16pt;
    font-weight: 700;
    color: #ffffff;
    background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
    text-align: center;
    padding: 20px 20px;
    margin: 30px 0 16px 0;
    letter-spacing: 1px;
    page-break-after: avoid;
  }}

  h2 {{
    font-size: 14pt;
    font-weight: 700;
    color: #0f3460;
    border-left: 6px solid #e94560;
    padding: 6px 0 6px 14px;
    margin-top: 28px;
    margin-bottom: 10px;
    background: #f0f4ff;
    page-break-after: avoid;
  }}

  h3 {{
    font-size: 12pt;
    font-weight: 700;
    color: #16213e;
    border-bottom: 2px solid #e94560;
    padding-bottom: 5px;
    margin-top: 22px;
    page-break-after: avoid;
  }}

  h4 {{
    font-size: 11pt;
    font-weight: 700;
    color: #e94560;
    margin-top: 16px;
    margin-bottom: 6px;
    page-break-after: avoid;
  }}

  p {{ margin: 6px 0 10px 0; }}

  strong {{ color: #0f3460; font-weight: 700; }}

  em {{ color: #555; font-style: italic; }}

  ul, ol {{
    margin: 6px 0 10px 22px;
    padding: 0;
  }}
  li {{ margin-bottom: 5px; }}

  /* TABLEAUX */
  table {{
    width: 100%;
    border-collapse: collapse;
    margin: 14px 0;
    font-size: 10pt;
    page-break-inside: avoid;
  }}
  th {{
    background: #0f3460;
    color: #fff;
    padding: 9px 11px;
    text-align: left;
    font-weight: 700;
    font-size: 10pt;
  }}
  td {{
    padding: 7px 11px;
    border-bottom: 1px solid #dde3f0;
    vertical-align: top;
  }}
  tr:nth-child(even) td {{ background: #f5f7fc; }}

  /* BLOCS CODE / PROMPTS */
  pre {{
    background: #1a1a2e;
    color: #a8f0e0;
    padding: 16px 18px;
    border-radius: 6px;
    border-left: 5px solid #e94560;
    font-family: 'Courier New', monospace;
    font-size: 9pt;
    white-space: pre-wrap;
    word-break: break-word;
    margin: 14px 0;
    page-break-inside: avoid;
  }}
  pre code {{
    background: none;
    color: inherit;
    padding: 0;
    font-size: inherit;
  }}

  /* CODE INLINE */
  code {{
    background: #eef1f9;
    color: #e94560;
    padding: 2px 6px;
    border-radius: 3px;
    font-family: 'Courier New', monospace;
    font-size: 9.5pt;
  }}

  /* CITATION */
  blockquote {{
    border-left: 5px solid #e94560;
    background: #fff5f7;
    margin: 14px 0;
    padding: 12px 18px;
    border-radius: 0 6px 6px 0;
    color: #333;
    font-style: italic;
  }}

  /* SÉPARATEUR */
  hr {{
    border: none;
    border-top: 2px solid #e0e6f5;
    margin: 24px 0;
  }}
</style>
</head>
<body>
{html_body}
</body>
</html>"""

output_path = "/home/user/fatou/TAMOU_NEURAL_PATH_Jour1_Lundi.pdf"
HTML(string=html, base_url="/").write_pdf(output_path)
print(f"PDF généré : {output_path}")
