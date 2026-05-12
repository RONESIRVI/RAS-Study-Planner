import os
from html2image import Html2Image
from datetime import datetime

# Directory for automation
BASE_DIR = "output"
os.makedirs(BASE_DIR, exist_ok=True)
hti = Html2Image(output_path=BASE_DIR)

def create_pillar_schedule_image(tasks_data):
    # Prepare Content
    classes_html = ""
    for idx in [0, 1]:
        t = tasks_data[idx].get('topic', '')
        if t and '[' not in t:
            if ':' in t:
                subj, top = t.split(':', 1)
                classes_html += f"<li><strong>{subj.strip()}</strong><br><small>Topic: {top.strip()}</small></li>"
            else:
                classes_html += f"<li>{t}</li>"
            
    revisions_html = ""
    revisions = tasks_data[2].get('revisions', [])
    for rev in revisions[:3]:
        revisions_html += f"<li><strong>{rev['subject']}</strong><br><small>{rev['topic']}</small></li>"

    pyq_topic = tasks_data[0].get('topic', 'Topic Pending')

    html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
        <meta charset="UTF-8">
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;700&family=Noto+Sans+Devanagari:wght@400;700&display=swap');
            body {{
                background: #0F0F1E;
                color: white;
                font-family: 'Poppins', 'Noto Sans Devanagari', sans-serif;
                margin: 0;
                padding: 40px;
                width: 1200px;
                height: 600px;
                overflow: hidden;
            }}
            .header {{
                font-size: 42px;
                font-weight: 700;
                margin-bottom: 50px;
                text-transform: uppercase;
                letter-spacing: 2px;
            }}
            .date {{
                font-size: 20px;
                color: #888;
                margin-top: -40px;
                margin-bottom: 40px;
            }}
            .container {{
                display: flex;
                gap: 20px;
            }}
            .card {{
                background: white;
                border-radius: 20px;
                flex: 1;
                height: 480px;
                color: black;
                display: flex;
                flex-direction: column;
                overflow: hidden;
                box-shadow: 0 10px 30px rgba(0,0,0,0.5);
            }}
            .card-header {{
                padding: 20px;
                text-align: center;
                font-weight: 700;
                font-size: 24px;
                margin-top: 60px;
            }}
            .classes-h {{ background: #4DB6AC; }}
            .revision-h {{ background: #F06292; }}
            .pyq-h {{ background: #FBC02D; }}
            .mock-h {{ background: #BA68C8; }}
            
            .icon {{
                font-size: 50px;
                text-align: center;
                margin-top: -140px;
            }}
            .content {{
                padding: 20px;
                flex-grow: 1;
            }}
            .content h3 {{
                font-size: 22px;
                margin-bottom: 15px;
            }}
            .content ul {{
                list-style: none;
                padding: 0;
                font-size: 16px;
            }}
            .content li {{
                margin-bottom: 12px;
                padding-left: 25px;
                position: relative;
            }}
            .content li::before {{
                content: '✅';
                position: absolute;
                left: 0;
                font-size: 14px;
            }}
            .footer {{
                position: absolute;
                bottom: 20px;
                left: 50%;
                transform: translateX(-50%);
                color: #555;
            }}
        </style>
    </head>
    <body>
        <div class="header">CORE PILLARS OF RAJASTHAN RAS MENTORSHIP</div>
        <div class="date">DATE: {datetime.now().strftime('%d %B %Y')}</div>
        <div class="container">
            <div class="card">
                <div class="icon">🏫</div>
                <div class="card-header classes-h">CLASSES</div>
                <div class="content">
                    <h3>CLASSES 1, 2, 3</h3>
                    <ul>{classes_html}</ul>
                </div>
            </div>
            <div class="card">
                <div class="icon">🔄</div>
                <div class="card-header revision-h">REVISION</div>
                <div class="content">
                    <h3>DAILY REVISION</h3>
                    <ul>{revisions_html}</ul>
                </div>
            </div>
            <div class="card">
                <div class="icon">🔍</div>
                <div class="card-header pyq-h">PYQ TEST</div>
                <div class="content">
                    <h3>1. PYQ test</h3>
                    <ul><li>Study: {pyq_topic}</li></ul>
                </div>
            </div>
            <div class="card">
                <div class="icon">⏱️</div>
                <div class="card-header mock-h">MOCK TESTS</div>
                <div class="content">
                    <h3>2. MCQ Test</h3>
                    <ul><li>(LAST Day Subjects:- Topics)</li></ul>
                </div>
            </div>
        </div>
        <div class="footer">Designing Your Path to RAJASTHAN RAS</div>
    </body>
    </html>
    """
    
    # Generate Image
    hti.screenshot(html_str=html_content, save_as='Pillar_Schedule.png', size=(1280, 720))
    return os.path.join(BASE_DIR, "Pillar_Schedule.png")
