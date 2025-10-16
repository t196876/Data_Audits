import pdfplumber
import pandas as pd
import re

# Step 1: Extract structured lines from PDF
pdf_path = "combined 18.pdf"
all_lines = []
with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            all_lines.extend(lines)

# Step 2: Parse structured data
people = {}
for i in range(len(all_lines)):
    line = all_lines[i]
    if 'HH' in line and 'CV' in line and 'NAME' in line and 'ADDR' in line:
        hhid = re.search(r'HH\s+(\d+)', line)
        cv = re.search(r'CV\s+(\S+)', line)
        name = re.search(r'NAME\s+(.+)', line)
        addr1 = re.search(r'ADDR\s+(.+)', line)
        addr2 = re.search(r'ADDR2\s+(.+)', line)

        key = hhid.group(1) if hhid else name.group(1).strip()
        if key not in people:
            people[key] = {
                'full_name': name.group(1).strip() if name else '',
                'household_id': hhid.group(1) if hhid else '',
                'creative_version': cv.group(1) if cv else '',
                'full_address': '',
                'slots': []
            }

        if addr1 and addr2:
            people[key]['full_address'] = f"{addr1.group(1).strip()},{addr2.group(1).strip()}"

    # Capture slot CPNs
    slot_match = re.search(r'SLOT\s+(\d+)\s+CPN\s+(\d+)', line)
    if slot_match:
        slot_num = int(slot_match.group(1))
        cpn = slot_match.group(2)
        if key in people:
            while len(people[key]['slots']) < slot_num:
                people[key]['slots'].append('')
            people[key]['slots'][slot_num - 1] = cpn

# Step 3: Convert to DataFrame
pdf_rows = []
for person in people.values():
    row = {
        'full_name': person['full_name'],
        'household_id': person['household_id'],
        'creative_version': person['creative_version'],
        'full_address': person['full_address']
    }
    for idx in range(12):
        row[f'Slot {idx+1}'] = person['slots'][idx] if idx < len(person['slots']) else ''
    pdf_rows.append(row)

df_pdf = pd.DataFrame(pdf_rows)

# Step 4: Save to Excel
df_pdf.to_excel("test3.xlsx", index=False)
