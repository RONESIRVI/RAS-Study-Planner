import os
from html2image import Html2Image
from datetime import datetime

# Directory for automation
BASE_DIR = "output"
os.makedirs(BASE_DIR, exist_ok=True)
if os.environ.get("GITHUB_ACTIONS"):
    hti = Html2Image(output_path=BASE_DIR, custom_flags=['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--headless'])
else:
    hti = Html2Image(output_path=BASE_DIR)

def create_pillar_schedule_image(tasks_data):
    # Prepare Content (Dynamic Loops)
    classes_html = ""
    for task in tasks_data:
        sub = task.get('subject', '')
        top = task.get('topic', '')
        # Only add if it's a real class, not a placeholder
        if sub and '[' not in sub and 'task' in task and 'CLASSES' in task['task']:
            classes_html += f"<li><strong>{sub}</strong><br><small>{top}</small></li>"
            
    revisions_html = ""
    # Smart Find: Search for the Revision task in the list
    for item in tasks_data:
        if item.get('task') == 'REVISION':
            revisions = item.get('revisions', [])
            for rev in revisions:
                revisions_html += f"<li><strong>{rev['subject']}</strong><br><small>{rev['topic']}</small></li>"
            break

    # Prepare PYQ Test Section (All classes)
    pyq_topics_html = ""
    for task in tasks_data:
        if 'CLASSES' in task.get('task', ''):
            sub = task.get('subject', '')
            top = task.get('topic', '')
            pyq_topics_html += f"<li>Study: {sub}: {top}</li>"

    html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
        <meta charset="UTF-8">
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;600&family=Noto+Sans+Devanagari:wght@400;700&display=swap');
            
            :root {{
                --bg-gradient: radial-gradient(circle at 10% 20%, rgb(15, 15, 30) 0%, rgb(8, 23, 44) 90.1%);
                --glass-bg: rgba(255, 255, 255, 0.05);
                --glass-border: rgba(255, 255, 255, 0.1);
                --accent-color: #00d2ff;
            }}

            body {{
                margin: 0;
                padding: 0;
                width: 1200px;
                min-height: 675px;
                background: var(--bg-gradient);
                /* Prioritize Noto, then fallback to Windows Standard Hindi Fonts */
                font-family: 'Outfit', 'Noto Sans Devanagari', 'Mangal', 'Arial Unicode MS', 'Nirmala UI', sans-serif;
                color: white;
                position: relative;
            }}

            .container {{
                padding: 40px;
            }}

            header {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 40px;
                border-bottom: 1px solid var(--glass-border);
                padding-bottom: 20px;
            }}

            .logo-area h1 {{
                font-size: 42px;
                margin: 0;
                background: linear-gradient(90deg, #fff, #00d2ff);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
            }}

            .grid {{
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 20px;
                align-items: start;
            }}

            .card {{
                background: var(--glass-bg);
                backdrop-filter: blur(10px);
                border: 1px solid var(--glass-border);
                border-radius: 20px;
                padding: 20px;
            }}

            .card-header {{
                font-size: 12px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: 2px;
                margin-bottom: 15px;
                padding: 5px 10px;
                border-radius: 5px;
                display: inline-block;
            }}

            .cls-h {{ background: #00d2ff22; color: #00d2ff; }}
            .rev-h {{ background: #ff009922; color: #ff0099; }}
            .pyq-h {{ background: #ffb80022; color: #ffb800; }}
            .mock-h {{ background: #c040ff22; color: #c040ff; }}

            h3 {{ font-size: 18px; margin: 10px 0; color: #fff; opacity: 0.9; }}
            ul {{ list-style: none; padding: 0; margin: 0; }}
            li {{
                margin-bottom: 12px;
                font-size: 14px;
                line-height: 1.4;
                color: rgba(255, 255, 255, 0.85);
                border-left: 2px solid var(--accent-color);
                padding-left: 10px;
            }}
            strong {{ display: block; margin-bottom: 2px; color: #fff; }}
            small {{ font-size: 12px; color: #00d2ff; opacity: 0.8; }}

            .footer-tag {{
                padding: 40px;
                text-align: right;
                font-size: 12px;
                color: rgba(255, 255, 255, 0.2);
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <header>
                <div class="logo-area"><h1>CORE PILLARS</h1></div>
                <div style="font-size: 18px; color: #00d2ff;">DATE: {datetime.now().strftime('%d %b %Y')}</div>
            </header>
            
            <div class="grid">
                <div class="card">
                    <div class="card-header cls-h">CLASSES</div>
                    <h3>DAILY FOCUS</h3>
                    <ul>{classes_html}</ul>
                </div>

                <div class="card">
                    <div class="card-header rev-h">REVISION</div>
                    <h3>SPACED REP</h3>
                    <ul>{revisions_html}</ul>
                </div>

                <div class="card">
                    <div class="card-header pyq-h">PYQ TEST</div>
                    <h3>ASSESSMENT</h3>
                    <ul>{pyq_topics_html}</ul>
                </div>

                <div class="card">
                    <div class="card-header mock-h">MOCK TEST</div>
                    <h3>STRENGTHEN</h3>
                    <ul>{pyq_topics_html}</ul>
                </div>
            </div>
            <div class="footer-tag">AIR-01 RAS MENTORSHIP | DYNAMIC ROADMAP</div>
        </div>
    </body>
    </html>
    """
    
    # Generate Image with dynamic height handling
    hti.screenshot(html_str=html_content, save_as='Pillar_Schedule.png', size=(1280, 800))
    return os.path.join(BASE_DIR, "Pillar_Schedule.png")
