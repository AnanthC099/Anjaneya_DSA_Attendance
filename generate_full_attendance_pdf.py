import csv
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

CSV_FILE = "Student_Attendance_Table.csv"
OUTPUT_FILE = "Full_Attendance_Students.pdf"
THRESHOLD_MINUTES = 180  # 3 hours

# Read CSV
with open(CSV_FILE, newline="") as f:
    reader = csv.reader(f)
    headers = next(reader)
    rows = list(reader)

# Filter: students who attended ALL sessions (no '--')
full_attendance = []
for row in rows:
    sessions = row[1:]
    if all(val.strip() != "--" for val in sessions):
        full_attendance.append(row)

session_headers = headers[1:]  # Session 36..43

# PDF setup
doc = SimpleDocTemplate(
    OUTPUT_FILE,
    pagesize=landscape(A4),
    leftMargin=15 * mm,
    rightMargin=15 * mm,
    topMargin=15 * mm,
    bottomMargin=15 * mm,
)

styles = getSampleStyleSheet()
elements = []

# Title
title_style = ParagraphStyle(
    "Title",
    parent=styles["Title"],
    fontSize=16,
    textColor=colors.HexColor("#1F4E79"),
    spaceAfter=4,
    alignment=TA_CENTER,
)
elements.append(Paragraph("Students Who Attended All Sessions (36-43)", title_style))

# Subtitle
subtitle_style = ParagraphStyle(
    "Subtitle",
    parent=styles["Normal"],
    fontSize=9,
    textColor=colors.HexColor("#996600"),
    alignment=TA_CENTER,
    fontName="Helvetica-Oblique",
    spaceAfter=10,
)
elements.append(
    Paragraph(
        f"Yellow = Attended less than 3 hours (&lt; {THRESHOLD_MINUTES} min)  |  Values are in minutes",
        subtitle_style,
    )
)

# Build table data
table_header = ["Sr.", "Student Name"] + [h.replace("Session ", "S") for h in session_headers]
table_data = [table_header]

for idx, row_data in enumerate(full_attendance, start=1):
    table_row = [str(idx), row_data[0]] + [val.strip() for val in row_data[1:]]
    table_data.append(table_row)

# Column widths
col_widths = [10 * mm, 60 * mm] + [22 * mm] * len(session_headers)

table = Table(table_data, colWidths=col_widths, repeatRows=1)

# Base table style
style_commands = [
    # Header
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTSIZE", (0, 0), (-1, 0), 8),
    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    # Data rows
    ("FONTNAME", (0, 1), (0, -1), "Helvetica"),
    ("FONTNAME", (1, 1), (1, -1), "Helvetica-Bold"),
    ("FONTNAME", (2, 1), (-1, -1), "Helvetica"),
    ("FONTSIZE", (0, 1), (-1, -1), 8),
    ("ALIGN", (0, 1), (0, -1), "CENTER"),
    ("ALIGN", (1, 1), (1, -1), "LEFT"),
    ("ALIGN", (2, 1), (-1, -1), "CENTER"),
    # Grid
    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#999999")),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#E6F0FA")]),
    # Padding
    ("TOPPADDING", (0, 0), (-1, -1), 2),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ("LEFTPADDING", (1, 1), (1, -1), 4),
]

# Yellow highlight for cells < 180 minutes
for row_idx, row_data in enumerate(full_attendance, start=1):
    for col_idx, val in enumerate(row_data[1:], start=2):
        minutes = int(val.strip())
        if minutes < THRESHOLD_MINUTES:
            style_commands.append(
                ("BACKGROUND", (col_idx, row_idx), (col_idx, row_idx), colors.yellow)
            )

table.setStyle(TableStyle(style_commands))
elements.append(table)

# Summary
elements.append(Spacer(1, 8 * mm))
summary_style = ParagraphStyle(
    "Summary", parent=styles["Normal"], fontSize=10, fontName="Helvetica-Bold"
)
elements.append(Paragraph(f"Total students with full attendance: {len(full_attendance)}", summary_style))
elements.append(Paragraph(f"Total sessions: {len(session_headers)}", summary_style))

doc.build(elements)
print(f"Created '{OUTPUT_FILE}' with {len(full_attendance)} students.")
