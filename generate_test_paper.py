import os
import re
import sys
import openpyxl
import docx
from docx.shared import Inches, Pt
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Ensure utf-8 terminal print support
sys.stdout.reconfigure(encoding='utf-8')

# Helper to add custom XML borders to table cells
def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key, val in edge_data.items():
                element.set(qn('w:{}'.format(key)), str(val))

# Smart splitting logic for bilingual question/option text
def split_bilingual(text):
    if not text:
        return "", ""
    text = str(text).strip()
    
    # Case 1: Split by double newline + bracket e.g. "Hindi \n\n[English]"
    match = re.search(r'^(.*?)\s*\n+\s*\[(.*?)\]\s*$', text, re.DOTALL)
    if match:
        return match.group(1).strip(), match.group(2).strip()
        
    # Case 2: Split by bracket at the end e.g. "Hindi [English]"
    if "[" in text and text.endswith("]"):
        parts = text.split("[")
        if len(parts) >= 2:
            hindi = "".join(parts[:-1]).strip()
            english = parts[-1].rstrip("]").strip()
            return hindi, english
            
    # Case 3: Split by " / " divider
    if " / " in text:
        parts = text.split(" / ")
        return parts[0].strip(), parts[1].strip()
        
    # Case 4: Split by "/" divider
    if "/" in text:
        parts = text.split("/")
        return parts[0].strip(), parts[1].strip()
        
    return text, ""

def generate_test_from_excel(excel_path, docx_path, sheet_name=None):
    print(f"Reading questions from Excel: {excel_path}")
    if not os.path.exists(excel_path):
        print(f"[ERROR] Excel file not found at: {excel_path}")
        return
        
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if sheet_name is None:
        # Default to the first sheet that matches 'PYQ Questions'
        sheet_name = next((s for s in wb.sheetnames if "PYQ Questions" in s), wb.sheetnames[0])
    
    ws = wb[sheet_name]
    print(f"Using sheet: {sheet_name}")
    
    # Extract questions data from rows
    questions = []
    # Identify column indices
    q_col, opt_start, ans_col = 3, 4, 8 # Defaults based on standard format
    
    # Read rows (skipping metadata rows)
    row_count = 0
    for r_idx, row in enumerate(ws.iter_rows(min_row=4, values_only=True), 4):
        if not row or len(row) < 9: continue
        q_val = row[q_col]
        if not q_val: continue
        
        # Options
        opts = [row[opt_start], row[opt_start+1], row[opt_start+2], row[opt_start+3]]
        
        questions.append({
            'qno': len(questions) + 1,
            'question': str(q_val),
            'options': [str(o) if o is not None else "" for o in opts],
        })
        row_count += 1

    print(f"Parsed {len(questions)} questions.")
    
    # Build docx Document
    doc = docx.Document()
    
    # 1. Cover Page Setup (Section 1)
    cover_section = doc.sections[0]
    cover_section.top_margin = Inches(0.6)
    cover_section.bottom_margin = Inches(0.6)
    cover_section.left_margin = Inches(0.6)
    cover_section.right_margin = Inches(0.6)
    
    # RPSC Header on cover page
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_h1 = p_header.add_run("राजस्थान राज्य एवं अधीनस्थ सेवाएँ संयुक्त (प्रा.) प्रतियोगी परीक्षा-2024\n")
    run_h1.font.name = 'Arial'
    run_h1.font.size = Pt(18)
    run_h1.font.bold = True
    
    run_h2 = p_header.add_run("परीक्षा दिनांक - 02.02.2025")
    run_h2.font.name = 'Arial'
    run_h2.font.size = Pt(14)
    run_h2.font.bold = True
    
    # Table for Booklet metadata
    table_meta = doc.add_table(rows=3, cols=3)
    table_meta.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table_meta.autofit = False
    
    col_widths = [Inches(2.3), Inches(2.7), Inches(2.3)]
    for r in table_meta.rows:
        for idx, w in enumerate(col_widths):
            r.cells[idx].width = w

    # Metadata cells content
    table_meta.cell(0, 0).paragraphs[0].add_run("पुस्तिका में पृष्ठों की संख्या : 40\n[Number of Pages in Booklet : 40]").font.size = Pt(9.5)
    
    c01 = table_meta.cell(0, 1).paragraphs[0]
    c01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_asr = c01.add_run("ASR-24")
    r_asr.font.size = Pt(20)
    r_asr.font.bold = True
    
    table_meta.cell(0, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    table_meta.cell(0, 2).paragraphs[0].add_run("प्रश्न-पुस्तिका संख्या व बारकोड\n[Question Booklet No. & Barcode]").font.size = Pt(9.5)

    table_meta.cell(1, 0).paragraphs[0].add_run("पुस्तिका में प्रश्नों की संख्या : 150\n[No. of Questions in Booklet : 150]").font.size = Pt(9.5)
    
    c11 = table_meta.cell(1, 1).paragraphs[0]
    c11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c11.add_run("इस प्रश्न-पुस्तिका को तब तक न खोलें जब तक कहा न जाए\n[Do not open this Booklet until asked]").font.size = Pt(9)
    
    table_meta.cell(1, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    table_meta.cell(1, 2).paragraphs[0].add_run("Paper Code : 00").font.size = Pt(11)
    
    table_meta.cell(2, 0).paragraphs[0].add_run("समय : 03 घण्टे + 10 मिनट अतिरिक्त*\n[Time : 03 Hours + 10 Minutes Extra*]").font.size = Pt(9.5)
    
    c21 = table_meta.cell(2, 1).paragraphs[0]
    c21.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = c21.add_run("Sub : G.K. & G.S.")
    r_sub.font.size = Pt(12)
    r_sub.font.bold = True
    
    table_meta.cell(2, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    table_meta.cell(2, 2).paragraphs[0].add_run("अधिकतम अंक : 200\n[Maximum Marks : 200]").font.size = Pt(9.5)

    # Style metadata cell borders
    border_style = {"sz": 6, "val": "single", "color": "000000"}
    for row in table_meta.rows:
        for cell in row.cells:
            set_cell_border(cell, top=border_style, bottom=border_style, left=border_style, right=border_style)
            
    doc.add_paragraph().paragraph_format.space_before = Pt(10)

    # Cover Page Instructions table
    table_inst = doc.add_table(rows=1, cols=2)
    table_inst.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table_inst.autofit = False
    table_inst.columns[0].width = Inches(3.6)
    table_inst.columns[1].width = Inches(3.6)
    
    # Left Cell: Hindi
    cell_hin = table_inst.cell(0, 0)
    p_hin = cell_hin.paragraphs[0]
    r_hi = p_hin.add_run("परीक्षार्थियों के लिए निर्देश\n")
    r_hi.font.bold = True
    r_hi.font.size = Pt(11)
    p_hin.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    instructions_hin = [
        "प्रत्येक प्रश्न के लिये एक विकल्प भरना अनिवार्य है।",
        "सभी प्रश्नों के अंक समान हैं।",
        "प्रत्येक प्रश्न का मात्र एक ही उत्तर दीजिए। एक से अधिक उत्तर देने की दशा में प्रश्न के उत्तर को गलत माना जाएगा।",
        "OMR उत्तर-पत्रक इस प्रश्न-पुस्तिका के अन्दर रखा है। जब आपको प्रश्न-पुस्तिका खोलने को कहा जाए, तो उत्तर-पत्रक निकालकर ध्यान से केवल नीले बॉल पॉइंट पेन से विवरण भरें।",
        "कृपया अपना रोल नम्बर ओ.एम.आर. उत्तर-पत्रक पर सावधानीपूर्वक सही भरें। गलत रोल नम्बर भरने पर परीक्षार्थी स्वयं उत्तरदायी होगा।",
        "ओ.एम.आर. उत्तर-पत्रक में करेक्शन पेन/व्हाइटनर/सफेदा का उपयोग निषिद्ध है।",
        "प्रत्येक गलत उत्तर के लिए प्रश्न अंक का 1/3 भाग काटा जायेगा। गलत उत्तर से तात्पर्य अशुद्ध उत्तर अथवा किसी भी प्रश्न के एक से अधिक उत्तर से है।",
        "प्रत्येक प्रश्न के पाँच विकल्प दिये गये हैं, जिन्हें क्रमशः 1, 2, 3, 4, 5 अंकित किया गया है। अभ्यर्थी को सही उत्तर निर्दिष्ट करते हुए उनमें से केवल एक गोले (बबल) को उत्तर-पत्रक पर नीले बॉल पॉइंट पेन से गहरा करना है।",
        "यदि आप प्रश्न का उत्तर नहीं देना चाहते हैं तो उत्तर-पत्रक में पाँचवें (5) विकल्प को गहरा करें। यदि पाँच में से कोई भी गोला गहरा नहीं किया जाता है, तो ऐसे प्रश्न के लिये प्रश्न अंक का 1/3 भाग काटा जायेगा।",
        "प्रश्न-पत्र हल करने के उपरांत अभ्यर्थी अनिवार्य रूप से ओ.एम.आर. उत्तर-पत्रक जांच लें कि समस्त प्रश्नों के लिये एक विकल्प (गोला) भर दिया गया है। इसके लिये निर्धारित समय से 10 मिनट का अतिरिक्त समय दिया गया है।"
    ]
    
    for idx, text in enumerate(instructions_hin, 1):
        p = cell_hin.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        r_num = p.add_run(f"{idx}. ")
        r_num.font.bold = True
        r_num.font.size = Pt(8.5)
        r_txt = p.add_run(text)
        r_txt.font.size = Pt(8.5)

    # Right Cell: English
    cell_eng = table_inst.cell(0, 1)
    p_eng = cell_eng.paragraphs[0]
    r_en = p_eng.add_run("INSTRUCTIONS FOR CANDIDATES\n")
    r_en.font.bold = True
    r_en.font.size = Pt(11)
    p_eng.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    instructions_eng = [
        "It is mandatory to fill one option for each question.",
        "All questions carry equal marks.",
        "Only one answer is to be given for each question. If more than one answers are marked, it would be treated as wrong answer.",
        "The OMR Answer Sheet is inside this Question Booklet. When you are directed to open the Question Booklet, take out the Answer Sheet and fill in the particulars carefully with Blue Ball Point Pen only.",
        "Please correctly fill your Roll Number in OMR Answer Sheet. Candidate will themself be responsible for filling wrong Roll No.",
        "Use of Correction Pen/Whitner in the OMR Answer Sheet is strictly forbidden.",
        "1/3 part of the mark(s) of each question will be deducted for each wrong answer. A wrong answer means an incorrect answer or more than one answers for any question.",
        "Each question has five options marked as 1, 2, 3, 4, 5. You have to darken only one circle (bubble) indicating the correct answer on the Answer Sheet using BLUE BALL POINT PEN.",
        "If you are not attempting a question then you have to darken the circle '5'. If none of the five circles is darkened, one third (1/3) part of the marks of question shall be deducted.",
        "After solving question paper, candidate must ascertain that he/she has darkened one of the circles (bubbles) for each of the questions. Extra time of 10 minutes beyond scheduled time, is provided for this."
    ]
    
    for idx, text in enumerate(instructions_eng, 1):
        p = cell_eng.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        r_num = p.add_run(f"{idx}. ")
        r_num.font.bold = True
        r_num.font.size = Pt(8.5)
        r_txt = p.add_run(text)
        r_txt.font.size = Pt(8.5)

    # Style instructions border lines
    set_cell_border(cell_hin, right={"sz": 6, "val": "single", "color": "000000"})
    set_cell_border(cell_eng, left={"sz": 6, "val": "single", "color": "000000"})
    for cell in (cell_hin, cell_eng):
        set_cell_border(cell, top={"sz": 12, "val": "single", "color": "000000"}, bottom={"sz": 12, "val": "single", "color": "000000"})
    set_cell_border(cell_hin, left={"sz": 12, "val": "single", "color": "000000"})
    set_cell_border(cell_eng, right={"sz": 12, "val": "single", "color": "000000"})

    # 2. Inside Pages Section Setup (Single-column page setup but with a 2-column layout grid inside)
    inside_section = doc.add_section(WD_SECTION.NEW_PAGE)
    inside_section.top_margin = Inches(0.6)
    inside_section.bottom_margin = Inches(0.6)
    inside_section.left_margin = Inches(0.6)
    inside_section.right_margin = Inches(0.6)
    
    # We will build a 2-column table grid for exact question alignment
    table_ques = doc.add_table(rows=0, cols=2)
    table_ques.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table_ques.autofit = False
    table_ques.columns[0].width = Inches(3.6)
    table_ques.columns[1].width = Inches(3.6)
    
    # Style separators
    border_middle = {"sz": 6, "val": "single", "color": "D3D3D3"} # light gray vertical line
    border_bottom = {"sz": 4, "val": "single", "color": "E0E0E0"} # subtle question divider
    
    for q in questions:
        row = table_ques.add_row()
        row.cells[0].width = Inches(3.6)
        row.cells[1].width = Inches(3.6)
        
        # Split bilingual content
        hindi_q, english_q = split_bilingual(q['question'])
        
        # Parse options
        hindi_opts, english_opts = [], []
        for o in q['options']:
            ho, eo = split_bilingual(o)
            hindi_opts.append(ho)
            english_opts.append(eo if eo else ho) # fallback to hindi if no translation
            
        # Add RPSC 5th option (Not Attempted)
        hindi_opts.append("अनुत्तरित प्रश्न")
        english_opts.append("Question not attempted")
        
        # Populate Hindi Cell (Left Column)
        cell_q_hin = row.cells[0]
        p_q_hin = cell_q_hin.paragraphs[0]
        p_q_hin.paragraph_format.space_after = Pt(6)
        p_q_hin.paragraph_format.line_spacing = 1.15
        
        # Question Number and Text
        r_num = p_q_hin.add_run(f"{q['qno']}.  ")
        r_num.font.bold = True
        r_num.font.size = Pt(9.5)
        
        r_text = p_q_hin.add_run(hindi_q)
        r_text.font.size = Pt(9.5)
        
        # Options Listing
        for o_idx, o_txt in enumerate(hindi_opts, 1):
            p_q_hin.add_run(f"\n    ({o_idx}) {o_txt}").font.size = Pt(9.5)
            
        # Populate English Cell (Right Column)
        cell_q_eng = row.cells[1]
        p_q_eng = cell_q_eng.paragraphs[0]
        p_q_eng.paragraph_format.space_after = Pt(6)
        p_q_eng.paragraph_format.line_spacing = 1.15
        
        # Question Number and Text
        r_num_e = p_q_eng.add_run(f"{q['qno']}.  ")
        r_num_e.font.bold = True
        r_num_e.font.size = Pt(9.5)
        
        r_text_e = p_q_eng.add_run(english_q if english_q else hindi_q)
        r_text_e.font.size = Pt(9.5)
        
        # Options Listing
        for o_idx, o_txt in enumerate(english_opts, 1):
            p_q_eng.add_run(f"\n    ({o_idx}) {o_txt}").font.size = Pt(9.5)
            
        # Apply border separating lines
        set_cell_border(cell_q_hin, right=border_middle, bottom=border_bottom)
        set_cell_border(cell_q_eng, left=border_middle, bottom=border_bottom)
        
    doc.save(docx_path)
    print(f"[SUCCESS] Professional Test Paper generated at: {docx_path}")

if __name__ == "__main__":
    import sys
    excel_p = r"R:\Final_PYQ_\Modal Qus Paper.xlsx"
    docx_p = r"R:\Final_PYQ_\Modal Qus Paper.docx"
    
    if len(sys.argv) > 1:
        excel_p = sys.argv[1]
    if len(sys.argv) > 2:
        docx_p = sys.argv[2]
        
    generate_test_from_excel(excel_p, docx_p)
