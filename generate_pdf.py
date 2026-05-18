import markdown
from weasyprint import HTML, CSS
from pathlib import Path

md_content = Path("/home/user/fatou/programme_IA_3mois_intensif.md").read_text(encoding="utf-8")

html_body = markdown.markdown(md_content, extensions=["tables", "fenced_code"])

html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');

  body {{
    font-family: 'Inter', Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.7;
    color: #1a1a2e;
    margin: 0;
    padding: 0;
  }}

  @page {{
    margin: 2.2cm 2cm 2cm 2cm;
    @bottom-center {{
      content: counter(page) " / " counter(pages);
      font-size: 9pt;
      color: #888;
    }}
  }}

  h1 {{
    font-size: 26pt;
    font-weight: 700;
    color: #ffffff;
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    text-align: center;
    padding: 40px 30px;
    margin: 0 0 24px 0;
    letter-spacing: 2px;
    page-break-after: avoid;
  }}

  h2 {{
    font-size: 15pt;
    font-weight: 700;
    color: #0f3460;
    border-left: 5px solid #e94560;
    padding-left: 12px;
    margin-top: 30px;
    margin-bottom: 10px;
    page-break-after: avoid;
  }}

  h3 {{
    font-size: 12.5pt;
    font-weight: 600;
    color: #16213e;
    border-bottom: 1px solid #d0d7f0;
    padding-bottom: 4px;
    margin-top: 20px;
    page-break-after: avoid;
  }}

  h4 {{
    font-size: 11pt;
    font-weight: 600;
    color: #e94560;
    margin-top: 14px;
    page-break-after: avoid;
  }}

  p {{
    margin: 6px 0 10px 0;
  }}

  strong {{
    color: #0f3460;
  }}

  ul, ol {{
    margin: 6px 0 10px 20px;
    padding: 0;
  }}

  li {{
    margin-bottom: 4px;
  }}

  table {{
    width: 100%;
    border-collapse: collapse;
    margin: 14px 0;
    font-size: 10pt;
    page-break-inside: avoid;
  }}

  th {{
    background-color: #0f3460;
    color: #ffffff;
    padding: 8px 10px;
    text-align: left;
    font-weight: 600;
  }}

  td {{
    padding: 7px 10px;
    border-bottom: 1px solid #dde3f0;
  }}

  tr:nth-child(even) td {{
    background-color: #f4f6fb;
  }}

  code {{
    background-color: #eef1f9;
    color: #e94560;
    padding: 2px 5px;
    border-radius: 3px;
    font-family: 'Courier New', monospace;
    font-size: 9.5pt;
  }}

  pre {{
    background-color: #1a1a2e;
    color: #a8dadc;
    padding: 14px 16px;
    border-radius: 6px;
    border-left: 4px solid #e94560;
    font-family: 'Courier New', monospace;
    font-size: 9pt;
    white-space: pre-wrap;
    word-break: break-word;
    margin: 12px 0;
    page-break-inside: avoid;
  }}

  pre code {{
    background: none;
    color: inherit;
    padding: 0;
  }}

  blockquote {{
    border-left: 4px solid #e94560;
    background: #f9f0f2;
    margin: 12px 0;
    padding: 10px 16px;
    border-radius: 0 6px 6px 0;
    color: #444;
    font-style: italic;
  }}

  hr {{
    border: none;
    border-top: 2px solid #e94560;
    margin: 24px 0;
  }}

  .page-break {{
    page-break-before: always;
  }}
</style>
</head>
<body>
{html_body}
</body>
</html>"""

output_path = "/home/user/fatou/TAMOU_NEURAL_PATH.pdf"
HTML(string=html, base_url="/").write_pdf(output_path)
print(f"PDF généré : {output_path}")
