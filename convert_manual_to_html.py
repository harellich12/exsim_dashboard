
import re

def slugify(text):
    # Basic slugify: lowercase, remove non-word chars (except space/hyphen), replace spaces with hyphens
    text = text.lower()
    text = re.sub(r'[^\w\s-]', '', text) 
    text = re.sub(r'[\s]+', '-', text)
    return text

def parse_markdown_to_html(md_text):
    html = """
    <html>
    <head>
        <style>
            body { font-family: Arial, sans-serif; line-height: 1.6; max-width: 800px; margin: 0 auto; padding: 20px; }
            h1 { color: #2c3e50; border-bottom: 2px solid #2c3e50; }
            h2 { color: #34495e; margin-top: 30px; border-bottom: 1px solid #ddd; }
            h3 { color: #16a085; margin-top: 20px; }
            table { border-collapse: collapse; width: 100%; margin: 20px 0; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            code { background-color: #f8f8f8; padding: 2px 5px; border-radius: 3px; font-family: Consolas, monospace; }
            pre { background-color: #f8f8f8; padding: 10px; border-radius: 5px; overflow-x: auto; }
            blockquote { border-left: 4px solid #3498db; margin: 0; padding-left: 15px; color: #555; background: #f0fbff; padding: 10px; }
            /* Mermaid styling handled by library, just ensure container is visible */
            .mermaid { margin: 20px 0; text-align: center; }
        </style>
        <script type="module">
            import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs';
            mermaid.initialize({ startOnLoad: true });
        </script>
    </head>
    <body>
    """
    
    lines = md_text.split('\n')
    in_table = False
    in_code_block = False
    is_mermaid = False
    list_stack = []
    
    for line in lines:
        line = line.rstrip()
        
        # Code Blocks
        if line.startswith('```'):
            if in_code_block:
                if is_mermaid: # Check flag set below
                     html += "</div>\n"
                else:
                     html += "</pre>\n"
                in_code_block = False
                is_mermaid = False
            else:
                lang = line.strip()[3:]
                if lang == 'mermaid':
                    html += '<div class="mermaid">\n'
                    is_mermaid = True
                else:
                    html += "<pre>"
                    is_mermaid = False
                in_code_block = True
            continue
            
        if in_code_block:
            if is_mermaid:
                html += line + "\n" # pass raw mermaid code
            else:
                html += line.replace('<', '&lt;').replace('>', '&gt;') + "\n"
            continue
            
        # Tables
        if line.strip().startswith('|'):
            if not in_table:
                html += "<table>"
                in_table = True
            
            # Check if it's a separator line
            if '---' in line:
                continue
                
            cells = [c.strip() for c in line.strip('|').split('|')]
            row_html = "<tr>"
            for cell in cells:
                # Basic inline formatting
                cell = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', cell)
                if in_table and html.strip().endswith('<table>'): # First row logic assumption (header)
                     row_html += f"<th>{cell}</th>"
                else:
                     row_html += f"<td>{cell}</td>"
            row_html += "</tr>"
            html += row_html
            continue
        elif in_table:
            html += "</table>\n"
            in_table = False

        # Headers
        if line.startswith('# '):
            text = line[2:].strip()
            slug = slugify(text)
            html += f'<h1 id="{slug}">{text}</h1>'
            continue
        if line.startswith('## '):
            text = line[3:].strip()
            slug = slugify(text)
            html += f'<h2 id="{slug}">{text}</h2>'
            continue
        if line.startswith('### '):
            text = line[4:].strip()
            slug = slugify(text)
            html += f'<h3 id="{slug}">{text}</h3>'
            continue
            
        # Lists (Basic)
        if line.strip().startswith('- ') or line.strip().startswith('* '):
            if not list_stack:
                html += "<ul>\n"
                list_stack.append('ul')
            # Handle item
            content = line.strip()[2:]
            content = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', content)
            # Handle Links in lists (for TOC)
            # Markdown link: [text](url) -> <a href="url">text</a>
            content = re.sub(r'\[(.*?)\]\((.*?)\)', r'<a href="\2">\1</a>', content)
            
            html += f"<li>{content}</li>"
            continue
        elif list_stack and line.strip() == '':
             # Loose check for end of list
             pass
        elif list_stack and not (line.strip().startswith('- ') or line.strip().startswith('* ')):
             html += "</ul>\n"
             list_stack.pop()

        # Horizontal Rule
        if line.strip() == '---':
            html += "<hr>"
            continue

        # Blockquotes / Alerts
        if line.strip().startswith('>'):
            content = line.strip()[1:].strip()
            if '[!IMPORTANT]' in content or '[!WARNING]' in content:
                continue # Skip the alert tag line
            content = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', content)
            html += f"<blockquote>{content}</blockquote>"
            continue

        # Normal Text
        if line.strip():
            content = line
            content = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', content)
            content = re.sub(r'`(.*?)`', r'<code>\1</code>', content)
            # Handle Links in normal text
            content = re.sub(r'\[(.*?)\]\((.*?)\)', r'<a href="\2">\1</a>', content)
            html += f"<p>{content}</p>"
            
    html += """
    </body>
    </html>
    """
    return html

if __name__ == "__main__":
    try:
        with open("USER_MANUAL.md", "r", encoding="utf-8") as f:
            md_content = f.read()
        
        html_content = parse_markdown_to_html(md_content)
        
        with open("USER_MANUAL.html", "w", encoding="utf-8") as f:
            f.write(html_content)
            
        print("Successfully created USER_MANUAL.html")
    except Exception as e:
        print(f"Error: {e}")
