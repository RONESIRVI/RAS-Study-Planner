import os
import shutil
from html2image import Html2Image
from datetime import datetime, timedelta

# Directory for automation
BASE_DIR = "output"
os.makedirs(BASE_DIR, exist_ok=True)
if os.environ.get("GITHUB_ACTIONS"):
    hti = Html2Image(output_path=BASE_DIR, custom_flags=['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--headless'])
else:
    hti = Html2Image(output_path=BASE_DIR)

def get_day_suffix(day):
    if 11 <= day <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")

def create_pillar_schedule_image(tasks_data, target_date=None):
    def get_compact_class(count):
        if count > 12:
            return "super-compact-3col"
        elif count > 7:
            return "super-compact"
        elif count > 4:
            return "compact"
        return ""

    # Use absolute file:// URL so Chrome headless can always load it
    logo_abs_path = os.path.abspath(os.path.join("assets", "Gemini_Generated_Image.png"))
    logo_file_url = "file:///" + logo_abs_path.replace("\\", "/")

    # Date formatting logic matching DATE.jpg
    if target_date is None:
        target_date = datetime.now() + timedelta(days=1)
    day_name = target_date.strftime('%A').upper()
    day_num = target_date.day
    suffix = get_day_suffix(day_num)
    month_name = target_date.strftime('%B').upper()
    year_name = target_date.strftime('%Y')

    day_blocks_html = "".join([f'<span class="letter-block">{char}</span>' for char in day_name])
    date_html = f'<span class="date-num">{day_num}</span><span class="date-suffix">{suffix}</span><span class="date-month">{month_name}</span>'

    # Check if target plan date is Saturday (meaning we are generating combined weekend plan)
    is_weekend_plan = (target_date.weekday() == 5)
    weekend_banner_html = ""
    if is_weekend_plan:
        sunday_date = target_date + timedelta(days=1)
        sat_num = target_date.day
        sun_num = sunday_date.day
        sat_suffix = get_day_suffix(sat_num)
        sun_suffix = get_day_suffix(sun_num)
        month_name = target_date.strftime('%B').upper()
        year_name = target_date.strftime('%Y')
        
        if target_date.strftime('%B') != sunday_date.strftime('%B'):
            sun_month = sunday_date.strftime('%B').upper()
            weekend_dates = f"{sat_num}{sat_suffix} {month_name} & {sun_num}{sun_suffix} {sun_month} {year_name}"
        else:
            weekend_dates = f"{sat_num}{sat_suffix} & {sun_num}{sun_suffix} {month_name} {year_name}"
            
        weekend_banner_html = f"""
        <div class="weekend-banner">
            <div class="announcing-badge-container">
                <span class="badge-line"></span>
                <div class="announcing-badge">ANNOUNCING</div>
                <span class="badge-line"></span>
            </div>
            <div class="weekend-title">
                <span class="title-week">Week</span><span class="title-end">end</span>
            </div>
            <div class="divider-curve-container">
                <div class="divider-curve"></div>
            </div>
            <div class="mentoring-programme">
                Saturday & Sunday - {weekend_dates}
            </div>
        </div>
        """

    # Prepare Content (Dynamic Loops)
    classes_html = ""
    classes_count = 0
    for task in tasks_data:
        sub = task.get('subject', '')
        top = task.get('topic', '')
        # Only add if it's a real class, not a placeholder
        if sub and '[' not in sub and 'task' in task and 'CLASSES' in task['task']:
            classes_count += 1
            classes_html += f"""
            <li>
                <div class="item-border"></div>
                <strong>{sub}</strong>
                <small>{top}</small>
                <div class="status-dot"></div>
            </li>
            """
            
    revisions_html = ""
    revisions_count = 0
    # Smart Find: Search for the Revision task in the list
    for item in tasks_data:
        if item.get('task') == 'REVISION':
            revisions = item.get('revisions', [])
            revisions_count = len(revisions)
            for rev in revisions:
                revisions_html += f"""
                <li>
                    <div class="item-border"></div>
                    <strong>{rev['subject']}</strong>
                    <small>{rev['topic']}</small>
                    <div class="status-dot"></div>
                </li>
                """
            break

    # Prepare PYQ Test Section (All classes)
    pyq_topics_html = ""
    pyq_count = 0
    for task in tasks_data:
        if 'CLASSES' in task.get('task', ''):
            sub = task.get('subject', '')
            top = task.get('topic', '')
            if sub and '[' not in sub:
                pyq_count += 1
                pyq_topics_html += f"""
                <li>
                    <div class="item-border"></div>
                    <strong>{sub}</strong>
                    <small>{top}</small>
                    <div class="status-dot"></div>
                </li>
                """

    # Prepare Mock Test Section (Only Spaced Repetition revisions, not Same Day Rev)
    mock_test_html = ""
    mock_count = 0
    for item in tasks_data:
        if item.get('task') == 'REVISION':
            revisions = item.get('revisions', [])
            mock_items = [r for r in revisions if "Same Day Rev" not in r.get('topic', '')]
            mock_count = len(mock_items)
            for rev in mock_items:
                mock_test_html += f"""
                <li>
                    <div class="item-border"></div>
                    <strong>{rev['subject']}</strong>
                    <small>{rev['topic']}</small>
                    <div class="status-dot"></div>
                </li>
                """
            break

    # Empty State Handlers
    if not classes_html:
        classes_html = """
        <li class="empty-state">
            <strong>No Classes Scheduled</strong>
            <small>Self study day</small>
        </li>
        """
    if not revisions_html:
        revisions_html = """
        <li class="empty-state">
            <strong>No Revisions Scheduled</strong>
            <small>Focus on daily classes</small>
        </li>
        """
    if not pyq_topics_html:
        pyq_topics_html = """
        <li class="empty-state">
            <strong>No Tests Scheduled</strong>
            <small>Self study day</small>
        </li>
        """
    if not mock_test_html:
        mock_test_html = """
        <li class="empty-state">
            <strong>No Tests Scheduled</strong>
            <small>Self study day</small>
        </li>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
        <meta charset="UTF-8">
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=Noto+Sans+Devanagari:wght@400;600;700&display=swap');
            
            :root {{
                --bg-gradient: linear-gradient(135deg, #0d0f1d 0%, #15103c 35%, #2a0f44 70%, #0c101d 100%);
                --glass-bg: rgba(255, 255, 255, 0.03);
                --glass-border: rgba(255, 255, 255, 0.08);
                
                --color-classes: #10b981;
                --color-revision: #ec4899;
                --color-pyq: #eab308;
                --color-mock: #a855f7;
            }}

            body {{
                margin: 0;
                padding: 0;
                width: 1280px;
                height: 720px;
                zoom: 6.0;
                background: var(--bg-gradient);
                font-family: 'Outfit', 'Noto Sans Devanagari', 'Mangal', 'Arial Unicode MS', 'Nirmala UI', sans-serif;
                color: #e2e8f0;
                overflow: hidden;
                box-sizing: border-box;
                position: relative;
            }}

            .container {{
                padding: 40px;
                height: 100%;
                box-sizing: border-box;
                display: flex;
                flex-direction: column;
                justify-content: space-between;
                position: relative;
            }}

            /* Ambient Glow Backgrounds */
            .glow-1 {{
                position: absolute;
                width: 500px;
                height: 500px;
                background: radial-gradient(circle, rgba(0, 242, 254, 0.06) 0%, rgba(0,0,0,0) 70%);
                top: -150px;
                left: -150px;
                z-index: 1;
            }}

            .glow-2 {{
                position: absolute;
                width: 600px;
                height: 600px;
                background: radial-gradient(circle, rgba(168, 85, 247, 0.04) 0%, rgba(0,0,0,0) 70%);
                bottom: -200px;
                right: -100px;
                z-index: 1;
            }}

            header {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-bottom: 1px solid var(--glass-border);
                padding-bottom: 24px;
                z-index: 10;
            }}

            .logo-area {{
                display: flex;
                align-items: center;
                gap: 16px;
            }}

            .logo-img {{
                width: 80px;
                height: 80px;
                object-fit: cover;
                border-radius: 20px;
                border: 2px solid rgba(0, 210, 255, 0.35);
                box-shadow: 0 6px 24px rgba(0, 242, 254, 0.25);
                background: rgba(255,255,255,0.04);
            }}

            .logo-text {{
                display: flex;
                flex-direction: column;
            }}

            .logo-text h1 {{
                font-size: 34px;
                font-weight: 700;
                letter-spacing: 2px;
                margin: 0;
                background: linear-gradient(135deg, #ffffff 0%, #a5b4fc 100%);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
            }}

            .logo-text p {{
                margin: 4px 0 0 0;
                font-size: 11px;
                font-weight: 600;
                letter-spacing: 4px;
                color: #6366f1;
                text-transform: uppercase;
            }}

            /* Custom Date Badge style matching DATE.jpg format with year at bottom */
            .date-badge {{
                background: rgba(17, 25, 40, 0.65);
                border: 2px solid rgba(255, 255, 255, 0.1);
                border-radius: 16px;
                padding: 10px 16px;
                display: flex;
                flex-direction: column;
                align-items: center;
                gap: 6px;
                backdrop-filter: blur(16px);
                box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.3);
                width: 190px;
                box-sizing: border-box;
            }}

            .day-blocks {{
                display: flex;
                gap: 3px;
                justify-content: center;
            }}

            .letter-block {{
                width: 17px;
                height: 22px;
                background: linear-gradient(180deg, #00d2ff 0%, #0066ff 100%);
                color: #ffffff;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 12px;
                font-weight: 700;
                border-radius: 3px;
                text-shadow: 0 1px 2px rgba(0,0,0,0.3);
                box-sizing: border-box;
            }}

            .date-row {{
                display: flex;
                align-items: baseline;
                justify-content: center;
                font-size: 17px;
                font-weight: 700;
                letter-spacing: 0.5px;
                color: #e2e8f0;
                margin-top: 2px;
            }}

            .date-num {{
                color: #ff0844; /* Red */
                font-size: 20px;
                margin-right: 1px;
            }}

            .date-suffix {{
                font-size: 10px;
                color: #94a3b8;
                vertical-align: super;
                margin-right: 5px;
            }}

            .date-month {{
                color: #a5b4fc; /* Light Indigo */
                text-transform: uppercase;
            }}

            .date-divider {{
                width: 100%;
                border: 0;
                border-top: 1px solid rgba(255, 255, 255, 0.12);
                margin: 2px 0;
            }}

            .year-row {{
                font-size: 14px;
                font-weight: 700;
                letter-spacing: 4px;
                color: #6366f1; /* Indigo matching the subheader */
                text-align: center;
                text-transform: uppercase;
                margin-top: 1px;
            }}

            /* Weekend Banner Styles matching photo_2026-06-01_14-03-21 layout */
            .weekend-banner {{
                margin-top: 15px;
                margin-bottom: 5px;
                display: flex;
                flex-direction: column;
                align-items: center;
                width: 100%;
                z-index: 10;
                box-sizing: border-box;
            }}

            .announcing-badge-container {{
                display: flex;
                align-items: center;
                justify-content: center;
                width: 100%;
                gap: 15px;
            }}

            .badge-line {{
                flex-grow: 1;
                height: 1px;
                background: linear-gradient(90deg, transparent, rgba(236, 72, 153, 0.4), transparent);
            }}

            .announcing-badge {{
                border: 1.5px solid #ec4899;
                border-radius: 8px;
                padding: 3px 18px;
                font-family: 'Outfit', sans-serif;
                font-size: 11px;
                font-weight: 700;
                letter-spacing: 3px;
                color: #ffffff;
                background: rgba(236, 72, 153, 0.1);
                box-shadow: 0 0 10px rgba(236, 72, 153, 0.15);
            }}

            .weekend-title {{
                font-size: 42px;
                font-weight: 800;
                letter-spacing: 1px;
                margin: 4px 0 2px 0;
                display: flex;
                justify-content: center;
                align-items: center;
                text-transform: none; /* Keep capitalization exactly as input */
            }}

            .title-week {{
                color: #00d2ff;
                text-shadow: 0 0 15px rgba(0, 210, 255, 0.25);
            }}

            .title-end {{
                color: #ec4899;
                text-shadow: 0 0 15px rgba(236, 72, 153, 0.25);
                margin-left: 2px;
            }}

            .divider-curve-container {{
                width: 70%;
                display: flex;
                justify-content: center;
                margin-bottom: 6px;
            }}

            .divider-curve {{
                width: 100%;
                height: 2px;
                background: linear-gradient(90deg, transparent, #ec4899, #00d2ff, transparent);
                border-radius: 50%;
            }}

            .mentoring-programme {{
                background: linear-gradient(90deg, rgba(30, 64, 175, 0.6) 0%, rgba(30, 58, 138, 0.8) 50%, rgba(30, 64, 175, 0.6) 100%);
                border: 1px solid rgba(59, 130, 246, 0.3);
                border-radius: 4px;
                padding: 5px 30px;
                font-size: 12px;
                font-weight: 700;
                letter-spacing: 2px;
                color: #ffffff;
                text-align: center;
                text-transform: uppercase;
                box-shadow: 0 4px 15px rgba(30, 64, 175, 0.2);
                width: 60%;
            }}

            .grid {{
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 24px;
                align-items: stretch;
                flex-grow: 1;
                margin-top: 30px;
                margin-bottom: 10px;
                z-index: 10;
            }}

            .card {{
                background: linear-gradient(135deg, rgba(255, 255, 255, 0.05) 0%, rgba(255, 255, 255, 0.02) 100%);
                backdrop-filter: blur(10px);
                border: 1px solid var(--glass-border);
                border-radius: 24px;
                padding: 24px;
                display: flex;
                flex-direction: column;
                transition: all 0.3s ease;
                box-sizing: border-box;
                position: relative;
                overflow: hidden;
            }}

            .card::before {{
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: linear-gradient(115deg, transparent 40%, rgba(255,255,255,0.08) 45%, rgba(255,255,255,0.12) 50%, rgba(255,255,255,0.08) 55%, transparent 60%);
                pointer-events: none;
            }}

            /* Custom Glow per card */
            .card-classes {{
                box-shadow: 0 10px 30px rgba(16, 185, 129, 0.03);
            }}
            .card-revision {{
                box-shadow: 0 10px 30px rgba(236, 72, 153, 0.03);
            }}
            .card-pyq {{
                box-shadow: 0 10px 30px rgba(234, 179, 8, 0.03);
            }}
            .card-mock {{
                box-shadow: 0 10px 30px rgba(168, 85, 247, 0.03);
            }}

            .card-header {{
                font-size: 11px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: 2.5px;
                margin-bottom: 20px;
                padding: 6px 14px;
                border-radius: 50px;
                display: inline-block;
                width: fit-content;
            }}

            .cls-h {{ background: rgba(16, 185, 129, 0.15); color: var(--color-classes); border: 1px solid rgba(16, 185, 129, 0.35); box-shadow: 0 0 12px rgba(16, 185, 129, 0.1); }}
            .rev-h {{ background: rgba(236, 72, 153, 0.15); color: var(--color-revision); border: 1px solid rgba(236, 72, 153, 0.35); box-shadow: 0 0 12px rgba(236, 72, 153, 0.1); }}
            .pyq-h {{ background: rgba(234, 179, 8, 0.15); color: var(--color-pyq); border: 1px solid rgba(234, 179, 8, 0.35); box-shadow: 0 0 12px rgba(234, 179, 8, 0.1); }}
            .mock-h {{ background: rgba(168, 85, 247, 0.15); color: var(--color-mock); border: 1px solid rgba(168, 85, 247, 0.35); box-shadow: 0 0 12px rgba(168, 85, 247, 0.1); }}

            .card h3 {{
                font-size: 18px;
                font-weight: 600;
                margin: 0 0 20px 0;
                color: #ffffff;
                letter-spacing: 0.5px;
            }}

            .card ul {{
                list-style: none;
                padding: 0;
                margin: 0;
                display: flex;
                flex-direction: column;
                gap: 12px;
                overflow: hidden;
                flex-grow: 1;
            }}

            /* Compact list items when many are present */
            .card ul.compact {{
                display: grid;
                grid-template-columns: repeat(2, 1fr);
                gap: 8px;
            }}
            .card ul.compact li {{
                padding: 8px 12px;
                padding-right: 28px;
                gap: 2px;
                border-radius: 8px;
            }}
            .card ul.compact li .item-border {{
                top: 8px;
                bottom: 8px;
            }}
            .card ul.compact li strong {{
                font-size: 11.5px;
            }}
            .card ul.compact li small {{
                font-size: 9.5px;
            }}

            .card ul.super-compact {{
                display: grid;
                grid-template-columns: repeat(2, 1fr);
                gap: 4px;
            }}
            .card ul.super-compact li {{
                padding: 4px 8px;
                padding-right: 20px;
                gap: 1px;
                border-radius: 6px;
            }}
            .card ul.super-compact li .item-border {{
                top: 4px;
                bottom: 4px;
            }}
            .card ul.super-compact li strong {{
                font-size: 9.5px;
            }}
            .card ul.super-compact li small {{
                font-size: 8px;
            }}
            .card ul.super-compact li .status-dot {{
                width: 5px;
                height: 5px;
                right: 6px;
            }}

            .card ul.super-compact-3col {{
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 4px;
            }}
            .card ul.super-compact-3col li {{
                padding: 4px 6px;
                padding-right: 18px;
                gap: 1px;
                border-radius: 6px;
            }}
            .card ul.super-compact-3col li .item-border {{
                top: 4px;
                bottom: 4px;
            }}
            .card ul.super-compact-3col li strong {{
                font-size: 8.5px;
            }}
            .card ul.super-compact-3col li small {{
                font-size: 7.5px;
            }}
            .card ul.super-compact-3col li .status-dot {{
                width: 4px;
                height: 4px;
                right: 4px;
            }}

            .card li {{
                background: rgba(13, 17, 28, 0.7);
                border: 1px solid rgba(255, 255, 255, 0.05);
                border-radius: 12px;
                padding: 14px 16px;
                padding-right: 36px; /* space for the dot */
                position: relative;
                box-sizing: border-box;
                display: flex;
                flex-direction: column;
                gap: 4px;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
            }}

            .status-dot {{
                position: absolute;
                right: 16px;
                top: 50%;
                transform: translateY(-50%);
                width: 8px;
                height: 8px;
                border-radius: 50%;
            }}

            .card-classes .status-dot {{ background: var(--color-classes); box-shadow: 0 0 8px var(--color-classes); }}
            .card-revision .status-dot {{ background: var(--color-revision); box-shadow: 0 0 8px var(--color-revision); }}
            .card-pyq .status-dot {{ background: var(--color-pyq); box-shadow: 0 0 8px var(--color-pyq); }}
            .card-mock .status-dot {{ background: var(--color-mock); box-shadow: 0 0 8px var(--color-mock); }}

            .item-border {{
                position: absolute;
                left: 0;
                top: 14px;
                bottom: 14px;
                width: 3px;
                border-radius: 0 4px 4px 0;
            }}

            .card-classes .item-border {{ background: var(--color-classes); }}
            .card-revision .item-border {{ background: var(--color-revision); }}
            .card-pyq .item-border {{ background: var(--color-pyq); }}
            .card-mock .item-border {{ background: var(--color-mock); }}

            .card li strong {{
                display: block;
                font-size: 14.5px;
                font-weight: 600;
                color: #ffffff;
                line-height: 1.4;
                padding-left: 6px;
            }}

            .card li small {{
                display: block;
                font-size: 12px;
                font-weight: 400;
                color: #94a3b8;
                line-height: 1.4;
                padding-left: 6px;
            }}

            .empty-state {{
                background: rgba(255, 255, 255, 0.01) !important;
                border: 1px dashed rgba(255, 255, 255, 0.08) !important;
                justify-content: center;
                align-items: center;
                text-align: center;
                padding: 24px !important;
            }}

            .empty-state strong {{
                color: #64748b !important;
                padding-left: 0 !important;
            }}

            .empty-state small {{
                color: #475569 !important;
                padding-left: 0 !important;
            }}

            footer {{
                text-align: center;
                font-size: 11px;
                letter-spacing: 3px;
                color: #94a3b8;
                text-transform: uppercase;
                margin-top: 15px;
                z-index: 10;
            }}

            .sparkle-star {{
                position: absolute;
                bottom: 25px;
                right: 40px;
                width: 24px;
                height: 24px;
                background: linear-gradient(135deg, #ffffff 0%, #a5b4fc 100%);
                clip-path: polygon(50% 0%, 61% 39%, 100% 50%, 61% 61%, 50% 100%, 39% 61%, 0% 50%, 39% 39%);
                opacity: 0.8;
                filter: drop-shadow(0 0 6px rgba(165, 180, 252, 0.6));
                z-index: 10;
            }}
        </style>
    </head>
    <body>
        <div class="glow-1"></div>
        <div class="glow-2"></div>
        
        <div class="container">
            <header>
                <div class="logo-area">
                    <img class="logo-img" src="{logo_file_url}" alt="Logo">
                    <div class="logo-text">
                        <h1>RAS MENTORSHIP</h1>
                    </div>
                </div>
                <div class="date-badge">
                    <div class="day-blocks">
                        {day_blocks_html}
                    </div>
                    <div class="date-row">
                        {date_html}
                    </div>
                    <hr class="date-divider">
                    <div class="year-row">
                        {year_name}
                    </div>
                </div>
            </header>
            
            {weekend_banner_html}
            
            <div class="grid">
                <div class="card card-classes">
                    <div class="card-header cls-h">CLASSES</div>
                    <h3>DAILY FOCUS</h3>
                    <ul class="{get_compact_class(classes_count)}">
                        {classes_html}
                    </ul>
                </div>
 
                <div class="card card-revision">
                    <div class="card-header rev-h">REVISION</div>
                    <h3>SPACED REP</h3>
                    <ul class="{get_compact_class(revisions_count)}">
                        {revisions_html}
                    </ul>
                </div>
 
                <div class="card card-pyq">
                    <div class="card-header pyq-h">PYQ TEST</div>
                    <h3>ASSESSMENT</h3>
                    <ul class="{get_compact_class(pyq_count)}">
                        {pyq_topics_html}
                    </ul>
                </div>
 
                <div class="card card-mock">
                    <div class="card-header mock-h">MOCK TEST</div>
                    <h3>STRENGTHEN</h3>
                    <ul class="{get_compact_class(mock_count)}">
                        {mock_test_html}
                    </ul>
                </div>
            </div>
            
            <footer>
                "Small daily improvements are the key to staggering longterm results."
            </footer>
            <div class="sparkle-star"></div>
        </div>
    </body>
    </html>
    """
    
    # Generate Image with dynamic height handling
    hti.screenshot(html_str=html_content, save_as='Pillar_Schedule.png', size=(7680, 4320))
    return os.path.join(BASE_DIR, "Pillar_Schedule.png")
