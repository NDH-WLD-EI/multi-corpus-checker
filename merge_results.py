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

    # Add header row
    headers = [
        (2, "Overall_Status"),
        (3, "Match_Count"),
        (4, "Consolidated_Notes"),
        (5, "Rajpal_Status"),
        (6, "Rajpal_Entry"),
        (7, "Hindwi_Status"),
        (8, "Hindwi_Card_Word"),
        (9, "Hindwi_Meaning"),
        (10, "HinKhoj_Status"),
        (11, "HinKhoj_Card_Word"),
        (12, "HinKhoj_Meaning"),
        (13, "Shabdkosh_Status"),
        (14, "Shabdkosh_H1_Word"),
        (15, "Shabdkosh_Meaning"),
    ]
    for col_idx, h_name in headers:
        ws.cell(row=1, column=col_idx, value=h_name)

    # Populate rows
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if r_idx == 1 or not row or not row[0]:
            continue
        w = str(row[0]).strip()
        data = all_records.get(w)
        if data:
            ws.cell(row=r_idx, column=2, value=data.get('Overall_Status', ''))
            ws.cell(row=r_idx, column=3, value=data.get('Match_Count', ''))
            ws.cell(row=r_idx, column=4, value=data.get('Consolidated_Notes', ''))
            ws.cell(row=r_idx, column=5, value=data.get('Rajpal_Status', ''))
            ws.cell(row=r_idx, column=6, value=data.get('Rajpal_Entry', ''))
            ws.cell(row=r_idx, column=7, value=data.get('Hindwi_Status', ''))
            ws.cell(row=r_idx, column=8, value=data.get('Hindwi_Card_Word', ''))
            ws.cell(row=r_idx, column=9, value=data.get('Hindwi_Meaning', ''))
            ws.cell(row=r_idx, column=10, value=data.get('HinKhoj_Status', ''))
            ws.cell(row=r_idx, column=11, value=data.get('HinKhoj_Card_Word', ''))
            ws.cell(row=r_idx, column=12, value=data.get('HinKhoj_Meaning', ''))
            ws.cell(row=r_idx, column=13, value=data.get('Shabdkosh_Status', ''))
            ws.cell(row=r_idx, column=14, value=data.get('Shabdkosh_H1_Word', ''))
            ws.cell(row=r_idx, column=15, value=data.get('Shabdkosh_Meaning', ''))

    wb.save(OUTPUT_EXCEL)
    print(f"\n[SUCCESS] Successfully merged all shards into: '{OUTPUT_EXCEL}'!")


if __name__ == '__main__':
    main()
