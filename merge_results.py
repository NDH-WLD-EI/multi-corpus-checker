import os
import sys
import glob
import csv
import openpyxl

INPUT_EXCEL = "Validated_content_added.xlsx"
OUTPUT_EXCEL = "Validated_content_added_multi_checked.xlsx"
SHEET_NAME = "Validated_words"


def main():
    csv_files = sorted(glob.glob("shard_*_results.csv"))
    print(f"Found {len(csv_files)} shard result files: {csv_files}")

    all_records = {}
    for cf in csv_files:
        print(f"Reading '{cf}'...")
        with open(cf, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                w = row.get('Word', '').strip()
                if w:
                    all_records[w] = row

    print(f"Total unique validated words collected: {len(all_records)}")

    if not os.path.exists(INPUT_EXCEL):
        print(f"Error: Base file '{INPUT_EXCEL}' not found!")
        sys.exit(1)

    print(f"Loading '{INPUT_EXCEL}' to generate enriched file...")
    wb = openpyxl.load_workbook(INPUT_EXCEL)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    # Create a fresh clean workbook for the output
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Validated_words"

    # Define headers
    headers = [
        "Word",
        "Overall_Status",
        "Mismatch_Type",
        "Match_Count",
        "Consolidated_Notes",
        "Rajpal_Status",
        "Rajpal_Entry",
        "Hindwi_Status",
        "Hindwi_Card_Word",
        "Hindwi_Meaning",
        "HinKhoj_Status",
        "HinKhoj_Card_Word",
        "HinKhoj_Meaning",
        "Shabdkosh_Status",
        "Shabdkosh_H1_Word",
        "Shabdkosh_Meaning"
    ]
    out_ws.append(headers)

    # Populate rows, dropping multi-word phrases with spaces
    dropped_space_phrases = 0
    written_rows = 0

    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if r_idx == 1 or not row or not row[0]:
            continue
        w = str(row[0]).strip()
        
        # User rule: remove phrases with spaces like 'राष्ट्रीय पक्षी'
        if ' ' in w:
            dropped_space_phrases += 1
            continue

        data = all_records.get(w, {})
        row_values = [
            w,
            data.get('Overall_Status', 'PENDING'),
            data.get('Mismatch_Type', ''),
            data.get('Match_Count', ''),
            data.get('Consolidated_Notes', ''),
            data.get('Rajpal_Status', ''),
            data.get('Rajpal_Entry', ''),
            data.get('Hindwi_Status', ''),
            data.get('Hindwi_Card_Word', ''),
            data.get('Hindwi_Meaning', ''),
            data.get('HinKhoj_Status', ''),
            data.get('HinKhoj_Card_Word', ''),
            data.get('HinKhoj_Meaning', ''),
            data.get('Shabdkosh_Status', ''),
            data.get('Shabdkosh_H1_Word', ''),
            data.get('Shabdkosh_Meaning', '')
        ]
        out_ws.append(row_values)
        written_rows += 1

    out_wb.save(OUTPUT_EXCEL)
    print(f"\n[SUCCESS] Successfully merged results into '{OUTPUT_EXCEL}'!")
    print(f"Total words written: {written_rows}")
    print(f"Multi-word phrases removed: {dropped_space_phrases}")


if __name__ == '__main__':
    main()
