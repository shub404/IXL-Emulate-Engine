from openpyxl import load_workbook
from groq import Groq
import csv
import os
import re
import time

client = Groq(api_key="{your grok api key}")

INPUT_FILE  = "ixl_grade3_questions.xlsx"
OUTPUT_FILE = "ixl_grade3_questions_updated.csv"

wb = load_workbook(INPUT_FILE)
ws = wb.active

headers = {cell.value: cell.column for cell in ws[1]}
col_q = headers["Question Text"]
col_o = headers["Question Options"]
col_a = headers["Correct Answer"]

header_row = [cell.value for cell in ws[1]]

# Resume: count rows already written to CSV (csv.reader handles embedded newlines in cells)
start_data_row = 2
if os.path.exists(OUTPUT_FILE):
    with open(OUTPUT_FILE, 'r', newline='', encoding='utf-8') as f:
        rows_written = sum(1 for _ in csv.reader(f)) - 1  # subtract header
    if rows_written > 0:
        start_data_row = 2 + rows_written
        print(f"Resuming from Excel row {start_data_row} (skipping {rows_written} already written rows)")

csv_mode = 'a' if start_data_row > 2 else 'w'
csv_file = open(OUTPUT_FILE, csv_mode, newline='', encoding='utf-8')
writer = csv.writer(csv_file)

if csv_mode == 'w':
    writer.writerow(header_row)

updates = 0

for row_idx in range(start_data_row, ws.max_row + 1):
    question_text = str(ws.cell(row=row_idx, column=col_q).value or "").strip()
    options_text  = str(ws.cell(row=row_idx, column=col_o).value or "").strip()
    answer_text   = str(ws.cell(row=row_idx, column=col_a).value or "").strip()

    row_data = [ws.cell(row=row_idx, column=col).value for col in range(1, len(header_row) + 1)]

    if not question_text:
        writer.writerow(row_data)
        csv_file.flush()
        continue

    prompt = f"""You are a warm, patient Grade 3 math teacher reading questions aloud to 8-year-old students in a classroom.
This output will be used to train a Text-to-Speech model, so it must sound exactly like natural spoken English — never like written text.

RULE 1 — Pauses and rhythm (for TTS prosody):
- Use commas for short pauses and ellipses (...) for longer, deliberate pauses.
- Pause before every number, math term, or key comparison so students have time to process: "There are... four equal parts."
- Pause between the problem setup and the final question being asked.
- Slow down on shape names, quantities, and operations — treat them as important beats.

RULE 2 — Short sentences with complete information:
- Every sentence must be 15 words or fewer. Break one long idea into 2 or 3 short sentences.
- Every number, shape, quantity, unit, condition, and math detail from the raw question MUST appear in your output — nothing can be dropped, summarized, or implied. If the raw question mentions 4 sides, say 4 sides. If it mentions a specific unit, say that unit.
- Spread details across multiple short sentences rather than packing them into one.

RULE 3 — Options format:
- Format the options cleanly as "Option A: [text]", "Option B: [text]", etc.

RULE 4 — Answer announcement:
- Reveal the answer with a warm lead-in, a pause, then the answer, then brief encouragement.
- Pattern: "[Warm phrase]! The correct answer is... [answer]. [Short encouragement]."
- The warm phrase must vary — use "Wonderful!", "That's right!", "Let's see...", "And the answer is" etc.
- The encouragement must vary — use "Great thinking!", "Well done!", "You're doing so well!", "Excellent work!" etc.
- Never write "Correct Answer:" as a label — it must sound like a teacher speaking.

RULE 5 — No non-speech symbols:
- Never use: slashes, asterisks, brackets, hyphens as bullets, or markdown. Only words and standard punctuation that sound natural when spoken aloud.

Raw Data to convert:
Question: {question_text}
Options: {options_text}
Answer: {answer_text}

Output exactly 3 parts separated by |~| with no extra text before or after:
[Spoken question] |~| [Spoken options] |~| [Spoken answer announcement]

Example:
Here are two circles. They are exactly the same size. One circle is cut into... four equal parts. The other circle is cut into... two equal parts. Which circle has the smaller equal pieces?|~|Option A: The circle divided into fourths\nOption B: The circle divided into halves|~|Wonderful! The correct answer is... the circle divided into fourths. Excellent thinking, everyone!
"""

    while True:
        try:
            response = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2,
                max_tokens=1024,
            )
            result_text = response.choices[0].message.content.strip()
            parts = result_text.split('|~|')

            if len(parts) == 3:
                row_data[col_q - 1] = parts[0].strip()
                row_data[col_o - 1] = parts[1].strip()
                row_data[col_a - 1] = parts[2].strip()
                updates += 1
                print(f"Successfully processed row {row_idx - 1} ({updates} updates)")
            else:
                print(f"Skipped row {row_idx - 1} — unexpected format: {result_text[:80]}")

            writer.writerow(row_data)
            csv_file.flush()
            time.sleep(3)
            break

        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                if "PerDay" in err_str or "per_day" in err_str.lower():
                    print(f"\n[DAILY QUOTA EXHAUSTED] Your free-tier daily limit is used up.")
                    print(f"Create a new Google Cloud project and generate a fresh API key, then update line 9.")
                    print(f"Progress saved up to row {row_idx - 2}. Script will exit now.")
                    csv_file.close()
                    exit(1)
                wait = 65
                match = re.search(r"retryDelay.*?'(\d+)s'", err_str)
                if match:
                    wait = int(match.group(1)) + 5
                print(f"Rate limit hit on row {row_idx - 1}. Waiting {wait}s before retry...")
                time.sleep(wait)
                # loop retries the same row
            else:
                print(f"Error processing row {row_idx - 1}: {e}")
                writer.writerow(row_data)  # write original only for non-rate-limit errors
                csv_file.flush()
                break

csv_file.close()
print(f"Processing complete. {updates} rows updated. File saved as {OUTPUT_FILE}")
