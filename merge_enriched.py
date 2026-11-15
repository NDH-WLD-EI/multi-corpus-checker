import os
import sys
import glob
import csv
import openpyxl

INPUT_EXCEL = "corpus_ready_for_enrichment.xlsx"
OUTPUT_EXCEL = "Validated_content_added_enriched_5sources.xlsx"

def main():
    csv_files = sorted(glob.glob("shard_*_enriched.csv"))
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

    print(f"Total unique enriched records collected: {len(all_records)}")

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Validated_words"

    headers = [
        "Word", "Valid_Nukta_Variant", "Overall_Status", "Mismatch_Type", "Match_Count", "Consolidated_Notes",
        "Rajpal_Status", "Rajpal_Entry",
        "Hindwi_Status", "Hindwi_Card_Word", "Hindwi_Meaning",
        "HinKhoj_Status", "HinKhoj_Card_Word", "HinKhoj_Meaning",
        "Shabdkosh_Status", "Shabdkosh_H1_Word", "Shabdkosh_Meaning",
        "Rekhta_Status", "Rekhta_Word", "Rekhta_Meaning"
    ]
    out_ws.append(headers)

    if not os.path.exists(INPUT_EXCEL):
        print(f"Base '{INPUT_EXCEL}' not found, saving collected records directly...")
        for w, data in all_records.items():
            row_vals = [data.get(h, '') for h in headers]
            out_ws.append(row_vals)
    else:
        wb_base = openpyxl.load_workbook(INPUT_EXCEL, read_only=True)
        ws_base = wb_base.active
        base_rows = list(ws_base.iter_rows(values_only=True))[1:]
        
        for r in base_rows:
            w = str(r[0]).strip()
            data = all_records.get(w, {})
            row_vals = [
                w,
                data.get("Valid_Nukta_Variant", r[1] if len(r)>1 else ""),
                data.get("Overall_Status", "PENDING"),
                data.get("Mismatch_Type", ""),
                data.get("Match_Count", ""),
                data.get("Consolidated_Notes", ""),
                data.get("Rajpal_Status", r[6] if len(r)>6 else ""),
                data.get("Rajpal_Entry", r[7] if len(r)>7 else ""),
                data.get("Hindwi_Status", r[8] if len(r)>8 else ""),
                data.get("Hindwi_Card_Word", r[9] if len(r)>9 else ""),
                data.get("Hindwi_Meaning", r[10] if len(r)>10 else ""),
                data.get("HinKhoj_Status", r[11] if len(r)>11 else ""),
                data.get("HinKhoj_Card_Word", r[12] if len(r)>12 else ""),
                data.get("HinKhoj_Meaning", r[13] if len(r)>13 else ""),
                data.get("Shabdkosh_Status", ""),
                data.get("Shabdkosh_H1_Word", ""),
                data.get("Shabdkosh_Meaning", ""),
                data.get("Rekhta_Status", ""),
                data.get("Rekhta_Word", ""),
                data.get("Rekhta_Meaning", "")
            ]
            out_ws.append(row_vals)

    out_wb.save(OUTPUT_EXCEL)
    print(f"\n[SUCCESS] Successfully merged into '{OUTPUT_EXCEL}'!")

if __name__ == '__main__':
    main()
