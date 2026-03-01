#!/usr/bin/env python3
"""
Generate attendance reports for DSA24 Sessions 36-43.

List_1.pdf: Students who attended ALL 8 sessions, with >=3 sessions over 3 hrs (180 min).
List_2.pdf: Students with exactly 2 leaves, with >=3 sessions over 3 hrs (180 min).
List_3.pdf: All remaining students.

Formatting:
- "--" where absent, duration in minutes where present
- Yellow highlight for cells with duration < 180 minutes (3 hrs)
"""

import csv
import os
from fpdf import FPDF

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SESSION_FILES = {
    "Session 36": os.path.join(BASE_DIR, "Session_36", "participants_87901948469_2026_01_31.csv"),
    "Session 37": os.path.join(BASE_DIR, "Session_37", "participants_89208110803_2026_02_01.csv"),
    "Session 38": os.path.join(BASE_DIR, "Session_38", "participants_89837366134_2026_02_07.csv"),
    "Session 39": os.path.join(BASE_DIR, "Session_39", "participants_88492557764_2026_02_08.csv"),
    "Session 40": os.path.join(BASE_DIR, "Session_40", "participants_84125620887_2026_02_14.csv"),
    "Session 41": os.path.join(BASE_DIR, "Session_41", "participants_84598698913_2026_02_15.csv"),
    "Session 42": os.path.join(BASE_DIR, "Session_42", "participants_83599210157_2026_02_22.csv"),
    "Session 43": os.path.join(BASE_DIR, "Session_43", "participants_87190343408_2026_02_28.csv"),
}

SESSION_NAMES = list(SESSION_FILES.keys())
THRESHOLD_MINUTES = 180  # 3 hours


def parse_csv(filepath):
    """Parse a Zoom participants CSV and return dict of {email: (name, duration)}."""
    participants = {}
    with open(filepath, "r", encoding="utf-8-sig") as f:
        lines = f.readlines()

    # Find the participant data header (line starting with "Name (original name)")
    data_start = None
    for i, line in enumerate(lines):
        if line.strip().startswith("Name (original name)"):
            data_start = i
            break

    if data_start is None:
        return participants

    reader = csv.DictReader(lines[data_start:])
    for row in reader:
        name = row.get("Name (original name)", "").strip()
        email = row.get("Email", "").strip().lower()
        duration = int(row.get("Total duration (minutes)", "0").strip())
        guest = row.get("Guest", "Yes").strip()

        # Skip the host (Guest=No)
        if guest == "No":
            continue

        if not email:
            continue

        # If same email appears multiple times, keep the higher duration
        if email in participants:
            existing_name, existing_dur = participants[email]
            participants[email] = (name if name.startswith("DSA") else existing_name, max(existing_dur, duration))
        else:
            participants[email] = (name, duration)

    return participants


def build_master_data():
    """Build master attendance data across all sessions."""
    # email -> {name: str, sessions: {session_name: duration_or_None}}
    master = {}

    for session_name, filepath in SESSION_FILES.items():
        participants = parse_csv(filepath)
        for email, (name, duration) in participants.items():
            if email not in master:
                master[email] = {"name": name, "sessions": {s: None for s in SESSION_NAMES}}
            # Prefer DSA-prefixed name
            if name.startswith("DSA") and not master[email]["name"].startswith("DSA"):
                master[email]["name"] = name
            master[email]["sessions"][session_name] = duration

    return master


def categorize_students(master):
    """Categorize students into List 1, 2, 3."""
    list1 = []  # All sessions attended, >=3 sessions > 180 min
    list2 = []  # Exactly 2 leaves, >=3 sessions > 180 min
    list3 = []  # Everyone else

    for email, data in sorted(master.items(), key=lambda x: x[1]["name"]):
        sessions = data["sessions"]
        name = data["name"]

        attended_sessions = sum(1 for v in sessions.values() if v is not None)
        absences = len(SESSION_NAMES) - attended_sessions

        # Count sessions with duration > 180 minutes
        sessions_over_3hrs = sum(1 for v in sessions.values() if v is not None and v >= THRESHOLD_MINUTES)

        if absences == 0:
            # Attended all sessions
            if sessions_over_3hrs >= 3:
                list1.append((email, data))
            else:
                list3.append((email, data))
        elif absences == 2:
            # Exactly 2 leaves
            if sessions_over_3hrs >= 3:
                list2.append((email, data))
            else:
                list3.append((email, data))
        else:
            list3.append((email, data))

    return list1, list2, list3


class AttendancePDF(FPDF):
    def __init__(self, title):
        super().__init__(orientation="L", unit="mm", format="A4")
        self.report_title = title

    def header(self):
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 10, self.report_title, align="C", new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")


def generate_pdf(students, filename, title):
    """Generate a PDF attendance report."""
    pdf = AttendancePDF(title)
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=20)

    # Column widths
    sr_w = 10
    name_w = 55
    session_w = 24
    total_w = sr_w + name_w + len(SESSION_NAMES) * session_w

    # Table header
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(sr_w, 8, "Sr", border=1, align="C", fill=True)
    pdf.cell(name_w, 8, "Student Name", border=1, align="C", fill=True)
    for sn in SESSION_NAMES:
        # Show as "S36", "S37" etc.
        short = "S" + sn.split()[-1]
        pdf.cell(session_w, 8, short, border=1, align="C", fill=True)
    pdf.ln()

    # Data rows
    pdf.set_font("Helvetica", "", 7)
    for idx, (email, data) in enumerate(students, 1):
        name = data["name"]
        sessions = data["sessions"]

        row_h = 7

        # Check if we need a new page
        if pdf.get_y() + row_h > pdf.h - 20:
            pdf.add_page()
            # Re-draw header
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_fill_color(200, 200, 200)
            pdf.cell(sr_w, 8, "Sr", border=1, align="C", fill=True)
            pdf.cell(name_w, 8, "Student Name", border=1, align="C", fill=True)
            for sn in SESSION_NAMES:
                short = "S" + sn.split()[-1]
                pdf.cell(session_w, 8, short, border=1, align="C", fill=True)
            pdf.ln()
            pdf.set_font("Helvetica", "", 7)

        # Sr No
        pdf.set_fill_color(255, 255, 255)
        pdf.cell(sr_w, row_h, str(idx), border=1, align="C")

        # Name
        pdf.cell(name_w, row_h, name[:35], border=1, align="L")

        # Session durations
        for sn in SESSION_NAMES:
            duration = sessions.get(sn)
            if duration is None:
                # Absent
                pdf.set_fill_color(255, 255, 255)
                pdf.cell(session_w, row_h, "--", border=1, align="C")
            else:
                if duration < THRESHOLD_MINUTES:
                    # Yellow highlight for < 3 hrs
                    pdf.set_fill_color(255, 255, 0)
                    pdf.cell(session_w, row_h, str(duration), border=1, align="C", fill=True)
                else:
                    pdf.set_fill_color(255, 255, 255)
                    pdf.cell(session_w, row_h, str(duration), border=1, align="C")

        pdf.ln()

    # Summary
    pdf.ln(5)
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(0, 8, f"Total Students: {len(students)}", new_x="LMARGIN", new_y="NEXT")

    output_path = os.path.join(BASE_DIR, filename)
    pdf.output(output_path)
    print(f"Generated: {output_path} ({len(students)} students)")


def main():
    master = build_master_data()
    print(f"Total unique students found: {len(master)}")

    list1, list2, list3 = categorize_students(master)

    print(f"\nList 1 (All sessions, >=3 sessions > 3hrs): {len(list1)} students")
    print(f"List 2 (Exactly 2 leaves, >=3 sessions > 3hrs): {len(list2)} students")
    print(f"List 3 (Remaining): {len(list3)} students")

    generate_pdf(
        list1,
        "List_1.pdf",
        "List 1: Students with Perfect Attendance (All 8 Sessions, >= 3 Sessions over 3 hrs)"
    )

    generate_pdf(
        list2,
        "List_2.pdf",
        "List 2: Students with Exactly 2 Leaves (>= 3 Sessions over 3 hrs)"
    )

    generate_pdf(
        list3,
        "List_3.pdf",
        "List 3: Remaining Students"
    )


if __name__ == "__main__":
    main()
