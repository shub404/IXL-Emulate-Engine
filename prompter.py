from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from google import genai
from google.genai import types
import shutil
import os
import time

client = genai.Client(api_key="YOUR API KEY")

generation_config = types.GenerateContentConfig(
    temperature=0.2,
    top_p=0.8,
    top_k=40,
)

INPUT_FILE = "ixl_grade2_questions.xlsx"
OUTPUT_FILE = "ixl_grade2_questions_updated.xlsx"

if not os.path.exists(OUTPUT_FILE):
    shutil.copy(INPUT_FILE, OUTPUT_FILE)

wb = load_workbook(OUTPUT_FILE)
ws = wb.active

headers = {cell.value: cell.column for cell in ws[1]}
col_q  = headers["Question Text"]
col_o  = headers["Question Options"]
col_a  = headers["Correct Answer"]

cell_font = Font(name="Calibri", size=11)
wrap_top  = Alignment(vertical="top", wrap_text=True)

def save_workbook():
    while True:
        try:
            wb.save(OUTPUT_FILE)
            return
        except PermissionError:
            print(f"\n[!] {OUTPUT_FILE} is open. Close it in Excel, then press ENTER to resume...")
            input()

updates = 0

for row_idx in range(2, ws.max_row + 1):
    question_text = str(ws.cell(row=row_idx, column=col_q).value or "").strip()
    options_text  = str(ws.cell(row=row_idx, column=col_o).value or "").strip()
    answer_text   = str(ws.cell(row=row_idx, column=col_a).value or "").strip()

    if not question_text:
        continue

    if options_text.startswith("Option A:"):
        print(f"Skipping row {row_idx - 1} — already processed.")
        continue

    prompt = f"""You are a professional Grade 1 math teacher dictating questions to 6-year-old students.
Your English is perfectly fluent, encouraging, and clear.
Convert the following raw data into natural, spoken English phrases as if you are speaking directly to the class.

Raw Data:
Question: {question_text}
Options: {options_text}
Answer: {answer_text}

Output Rules:
1. Make the question sound engaging but mathematically precise for a 6-year-old.
2. Format the options cleanly as "Option A: [text]", "Option B: [text]", etc.
3. Format the answer as "Correct Answer: [text]".
4. You MUST separate the three generated parts using exactly this delimiter: |~|
5. Do not output any markdown formatting, conversational filler, or extra text.

Example Output Structure:
Let's look at two circles. They are the exact same size. One circle is divided into four equal parts, and the other is divided into two equal parts. Which circle has the smaller equal parts?|~|Option A: The circle divided into fourths\nOption B: The circle divided into halves|~|Correct Answer: The circle divided into fourths
"""

    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash", contents=prompt, config=generation_config
        )
        result_text = response.text.strip()
        parts = result_text.split('|~|')

        if len(parts) == 3:
            for col_idx, value in [(col_q, parts[0].strip()),
                                   (col_o, parts[1].strip()),
                                   (col_a, parts[2].strip())]:
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = cell_font
                cell.alignment = wrap_top
            updates += 1
            print(f"Successfully processed row {row_idx - 1} ({updates} updates)")
            if updates % 100 == 0:
                save_workbook()
                print(f"[Saved] {updates} updates written to {OUTPUT_FILE}")
        else:
            print(f"Skipped row {row_idx - 1} — unexpected format: {result_text[:80]}")

        time.sleep(3)

    except Exception as e:
        print(f"Error processing row {row_idx - 1}: {e}")
        continue

save_workbook()
print(f"Processing complete. {updates} rows updated. File saved as {OUTPUT_FILE}")
