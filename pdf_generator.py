import os
import sqlite3
import datetime
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, BaseDocTemplate, PageTemplate, Frame,
    Paragraph, Spacer, Table, TableStyle, NextPageTemplate,
    PageBreak, Image, ListFlowable, ListItem, Flowable, KeepTogether
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

DB_PATH = '/workspace/ndl_monitoring.db'

# Define Palette
COLORS = {
    'heading':    HexColor('#1F4E78'),
    'body':       HexColor('#2C2C2C'),
    'accent':     HexColor('#0070C0'),
    'muted':      HexColor('#7F7F7F'),
    'bg_alt':     HexColor('#F2F4F7'),
    'bg_header':  HexColor('#2F5496'),
    'white':      HexColor('#ffffff'),
    'present_bg': HexColor('#E2EFDA'),
    'present_fg': HexColor('#375623'),
    'absent_bg':  HexColor('#FCE4D6'),
    'absent_fg':  HexColor('#C65911'),
}

HEADING_FONT = 'Helvetica-Bold'
BODY_FONT    = 'Helvetica'
MONO_FONT    = 'Courier'

class SectionDivider(Flowable):
    def __init__(self, width, color):
        Flowable.__init__(self)
        self._width = width
        self.color = color
        self._height = 15

    def wrap(self, availWidth, availHeight):
        return self._width, self._height

    def draw(self):
        y = self._height / 2
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(1)
        self.canv.line(0, y, self._width, y)

def on_later_pages(canvas, doc):
    canvas.saveState()
    
    # Header rule
    canvas.setStrokeColor(COLORS['accent'])
    canvas.setLineWidth(0.6)
    y_rule = LETTER[1] - doc.topMargin + 14
    canvas.line(doc.leftMargin, y_rule, LETTER[0] - doc.rightMargin, y_rule)
    
    # Header text
    canvas.setFont(HEADING_FONT, 8)
    canvas.setFillColor(COLORS['muted'])
    canvas.drawString(doc.leftMargin, y_rule + 4, "NDL LEARNER MONITORING & ROLL-CALL APP")
    canvas.drawRightString(LETTER[0] - doc.rightMargin, y_rule + 4, "LIVE STATUS REPORT")
    
    # Footer
    y_footer = doc.bottomMargin - 24
    canvas.setStrokeColor(COLORS['bg_alt'])
    canvas.setLineWidth(0.3)
    canvas.line(doc.leftMargin, y_footer + 14, LETTER[0] - doc.rightMargin, y_footer + 14)
    canvas.setFont(BODY_FONT, 8)
    canvas.setFillColor(COLORS['muted'])
    canvas.drawString(doc.leftMargin, y_footer, f"Generated: {datetime.date.today().strftime('%Y-%m-%d')} | Shared Database Live Export")
    canvas.drawCentredString(LETTER[0] / 2, y_footer, f"Page {doc.page}")
    
    canvas.restoreState()

def on_first_page(canvas, doc):
    canvas.saveState()
    # Bottom rule for the footer on page 1
    y_footer = doc.bottomMargin - 24
    canvas.setStrokeColor(COLORS['bg_alt'])
    canvas.setLineWidth(0.3)
    canvas.line(doc.leftMargin, y_footer + 14, LETTER[0] - doc.rightMargin, y_footer + 14)
    canvas.setFont(BODY_FONT, 8)
    canvas.setFillColor(COLORS['muted'])
    canvas.drawString(doc.leftMargin, y_footer, f"Generated: {datetime.date.today().strftime('%Y-%m-%d')} | NDL Real-Time App")
    canvas.drawCentredString(LETTER[0] / 2, y_footer, f"Page {doc.page}")
    canvas.restoreState()

def generate_pdf_report(output_path):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Fetch learners and latest status
    cursor.execute('''
    SELECT 
        l.name, 
        l.assigned_teacher, 
        l.room,
        u.status,
        u.event,
        u.recorded_by,
        u.date,
        u.time
    FROM learners l
    LEFT JOIN (
        SELECT u1.* FROM updates u1
        JOIN (
            SELECT learner_id, MAX(timestamp) as max_ts 
            FROM updates 
            GROUP BY learner_id
        ) u2 ON u1.learner_id = u2.learner_id AND u1.timestamp = u2.max_ts
    ) u ON l.id = u.learner_id
    ORDER BY l.name ASC
    ''')
    learners_status = cursor.fetchall()
    
    # Get current statistics
    cursor.execute("SELECT COUNT(*) FROM learners")
    total_count = cursor.fetchone()[0]
    
    # Calculate counts manually
    present_count = 0
    absent_count = 0
    no_update_count = 0
    for l in learners_status:
        st = l[3]
        if not st:
            no_update_count += 1
        elif st == "Present":
            present_count += 1
        elif st == "Absent":
            absent_count += 1
            
    conn.close()
    
    # Initialize Document
    doc = BaseDocTemplate(
        output_path,
        pagesize=LETTER,
        leftMargin=0.5 * inch, rightMargin=0.5 * inch,
        topMargin=0.75 * inch, bottomMargin=0.75 * inch,
    )
    
    PAGE_W, PAGE_H = LETTER
    USABLE_W = PAGE_W - 2 * doc.leftMargin
    
    content_frame = Frame(
        doc.leftMargin, doc.bottomMargin,
        USABLE_W, PAGE_H - doc.topMargin - doc.bottomMargin,
        id='main'
    )
    
    doc.addPageTemplates([
        PageTemplate(id='first_page', frames=content_frame, onPage=on_first_page),
        PageTemplate(id='content', frames=content_frame, onPage=on_later_pages),
    ])
    
    styles = getSampleStyleSheet()
    
    # Custom Paragraph Styles
    style_title = ParagraphStyle(
        'DocTitle', fontName=HEADING_FONT, fontSize=18,
        textColor=COLORS['heading'], leading=22,
        spaceAfter=4, alignment=TA_LEFT
    )
    style_subtitle = ParagraphStyle(
        'DocSubtitle', fontName=BODY_FONT, fontSize=10,
        textColor=COLORS['muted'], leading=12,
        spaceAfter=12, alignment=TA_LEFT
    )
    style_h1 = ParagraphStyle(
        'H1', fontName=HEADING_FONT, fontSize=12,
        textColor=COLORS['heading'], leading=15,
        spaceBefore=12, spaceAfter=6,
    )
    style_body = ParagraphStyle(
        'Body', fontName=BODY_FONT, fontSize=9,
        textColor=COLORS['body'], leading=12,
        spaceAfter=4
    )
    style_table_head = ParagraphStyle(
        'TableHead', fontName=HEADING_FONT, fontSize=8,
        textColor=COLORS['white'], leading=10,
    )
    style_table_body = ParagraphStyle(
        'TableBody', fontName=BODY_FONT, fontSize=8,
        textColor=COLORS['body'], leading=10,
    )
    style_table_body_bold = ParagraphStyle(
        'TableBodyBold', fontName=HEADING_FONT, fontSize=8,
        textColor=COLORS['body'], leading=10,
    )
    style_metric_label = ParagraphStyle(
        'MetricLabel', fontName=HEADING_FONT, fontSize=9,
        textColor=COLORS['body'], leading=11,
    )
    style_metric_val = ParagraphStyle(
        'MetricVal', fontName=HEADING_FONT, fontSize=11,
        textColor=COLORS['heading'], leading=13,
        alignment=TA_CENTER
    )
    
    story = []
    
    # 1. Header block
    story.append(Paragraph("NDL Learner Monitoring & Roll-Call App", style_title))
    story.append(Paragraph(f"Live Safety Status Report — Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", style_subtitle))
    story.append(SectionDivider(USABLE_W, COLORS['accent']))
    story.append(Spacer(1, 4))
    
    # 2. Metrics Block (Side-by-Side as a table)
    metric_headers = [
        Paragraph("Total Learners", style_table_head),
        Paragraph("Present", style_table_head),
        Paragraph("Absent", style_table_head),
        Paragraph("No Update Yet", style_table_head)
    ]
    metric_values = [
        Paragraph(str(total_count), style_metric_val),
        Paragraph(str(present_count), style_metric_val),
        Paragraph(str(absent_count), style_metric_val),
        Paragraph(str(no_update_count), style_metric_val)
    ]
    
    m_table = Table([metric_headers, metric_values], colWidths=[USABLE_W/4.0]*4)
    m_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COLORS['bg_header']),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 1), (0, 1), COLORS['bg_alt']),
        ('BACKGROUND', (1, 1), (1, 1), COLORS['present_bg']),
        ('BACKGROUND', (2, 1), (2, 1), COLORS['absent_bg']),
        ('BACKGROUND', (3, 1), (3, 1), COLORS['bg_alt']),
        ('GRID', (0, 0), (-1, -1), 0.5, COLORS['muted']),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(m_table)
    story.append(Spacer(1, 10))
    
    # 3. Learner Status List
    story.append(Paragraph("Current Learner Status Overview", style_h1))
    
    headers = [
        Paragraph("Learner Name", style_table_head),
        Paragraph("Assigned Teacher", style_table_head),
        Paragraph("Room", style_table_head),
        Paragraph("Latest Status", style_table_head),
        Paragraph("Latest Event", style_table_head),
        Paragraph("Recorded By", style_table_head),
        Paragraph("Date / Time", style_table_head)
    ]
    
    rows = []
    for data in learners_status:
        name, teacher, room, status, event, recorded_by, d_val, t_val = data
        
        # Handle N/A displays
        room_disp = room if room else "Not Assigned"
        event_disp = event if event else "N/A"
        rec_disp = recorded_by if recorded_by else "N/A"
        dt_disp = f"{d_val} {t_val}" if d_val else "N/A"
        
        # Color coding status
        if not status:
            status_para = Paragraph("No update yet", style_table_body)
            row_bg_color = COLORS['white']
        elif status == "Present":
            status_para = Paragraph("Present", style_table_body_bold)
            row_bg_color = COLORS['present_bg']
        else:
            status_para = Paragraph("Absent", style_table_body_bold)
            row_bg_color = COLORS['absent_bg']
            
        rows.append([
            Paragraph(name, style_table_body_bold),
            Paragraph(teacher, style_table_body),
            Paragraph(room_disp, style_table_body),
            status_para,
            Paragraph(event_disp, style_table_body),
            Paragraph(rec_disp, style_table_body),
            Paragraph(dt_disp, style_table_body)
        ])
        
    col_widths = [USABLE_W * 0.25, USABLE_W * 0.15, USABLE_W * 0.08, USABLE_W * 0.12, USABLE_W * 0.13, USABLE_W * 0.13, USABLE_W * 0.14]
    
    t = Table([headers] + rows, colWidths=col_widths, repeatRows=1)
    
    t_styles = [
        ('BACKGROUND', (0, 0), (-1, 0), COLORS['bg_header']),
        ('GRID', (0, 0), (-1, -1), 0.5, COLORS['muted']),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]
    
    # Highlight status colors on individual rows
    for r_idx, r_data in enumerate(learners_status, 1):
        st = r_data[3]
        if st == "Present":
            t_styles.append(('BACKGROUND', (3, r_idx), (3, r_idx), COLORS['present_bg']))
            t_styles.append(('TEXTCOLOR', (3, r_idx), (3, r_idx), COLORS['present_fg']))
        elif st == "Absent":
            t_styles.append(('BACKGROUND', (3, r_idx), (3, r_idx), COLORS['absent_bg']))
            t_styles.append(('TEXTCOLOR', (3, r_idx), (3, r_idx), COLORS['absent_fg']))
            
        # Standard zebra row backgrounds for non-highlighted columns
        zebra_color = COLORS['bg_alt'] if r_idx % 2 == 0 else COLORS['white']
        for c_idx in [0, 1, 2, 4, 5, 6]:
            t_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), zebra_color))
            
    t.setStyle(TableStyle(t_styles))
    story.append(t)
    
    # We transitions to content template for multi-page document
    story.insert(0, NextPageTemplate('content'))
    
    doc.build(story)
