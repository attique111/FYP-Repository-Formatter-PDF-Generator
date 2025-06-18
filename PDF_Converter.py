import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import os

# === CONFIG ===
excel_file = "Fall-2021.xlsx"
pdf_file = "Fall-2021.pdf"
page_width, page_height = A4
top_margin = 50
bottom_margin = 70  # Increased from 50 to avoid cutting
side_margin = 40
line_height = 15

# === LOAD EXCEL ===
try:
    df = pd.read_excel(excel_file)
    df.columns = [col.strip() for col in df.columns]
    print("Loaded Excel. Columns found:", df.columns.tolist())
except Exception as e:
    print(f"❌ Failed to load Excel file: {e}")
    exit()

# Column mappings
title_col = "Unnamed: 1"
desc_col = "Unnamed: 2"

# === INIT PDF ===
c = canvas.Canvas(pdf_file, pagesize=A4)
y_position = page_height - top_margin

# === HEADER ===
c.setFont("Helvetica-Bold", 16)
c.drawString(side_margin, y_position, "FYP Ideas - Fall 2023")
y_position -= 30

# === WRITE EACH ROW ===
entry_number = 1

for index, row in df.iterrows():
    title = str(row.get(title_col, "")).strip()
    description = str(row.get(desc_col, "")).strip()

    if not title and not description:
        continue

    # Estimate lines needed for title and description
    def get_wrapped_lines(text, max_chars=110):
        lines = []
        for line in text.split("\n"):
            while len(line) > max_chars:
                split_idx = line[:max_chars].rfind(" ")
                if split_idx == -1:
                    split_idx = max_chars
                lines.append(line[:split_idx])
                line = line[split_idx:].lstrip()
            lines.append(line)
        return lines

    title_lines = get_wrapped_lines(f"{entry_number}. {title}")
    desc_lines = get_wrapped_lines(description)
    total_lines_needed = len(title_lines) + len(desc_lines) + 3  # including gaps

    # Check if there's space on the current page
    if y_position - (total_lines_needed * line_height) < bottom_margin:
        c.showPage()
        y_position = page_height - top_margin

    # --- Write Title ---
    c.setFont("Helvetica-Bold", 12)
    for line in title_lines:
        c.drawString(side_margin, y_position, line)
        y_position -= line_height

    y_position -= 5  # small gap

    # --- Write Description ---
    c.setFont("Helvetica", 11)
    for line in desc_lines:
        c.drawString(side_margin, y_position, line)
        y_position -= line_height

    y_position -= 15  # extra space between ideas
    entry_number += 1

# === SAVE PDF ===
c.save()
print(f"PDF generated successfully: {os.path.abspath(pdf_file)}")