import openpyxl
import sys
import unicodedata
import os
import shutil

# Force UTF-8 on Windows console
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

INPUT_EXCEL = r"C:\Users\nidhi.waldia.EI-LAP-8913\Downloads\downloads\Antigravity\check corpus\Recheck\Validated_content_added_multi_checked_repaired.xlsx"
OUTPUT_EXCEL = r"C:\Users\nidhi.waldia.EI-LAP-8913\Downloads\downloads\Antigravity\check corpus\multi-corpus-checker\Validated_content_added_nukta_standardized.xlsx"
RECHECK_DEST = r"C:\Users\nidhi.waldia.EI-LAP-8913\Downloads\downloads\Antigravity\check corpus\Recheck\Validated_content_added_nukta_standardized.xlsx"

def strip_nukta(text):
    decomposed = unicodedata.normalize('NFD', text)
    stripped = ''.join(c for c in decomposed if c != '\u093c')
    return unicodedata.normalize('NFC', stripped)

def has_nukta(text):
    return '\u093c' in unicodedata.normalize('NFD', text)

def main():
    print(f"Loading '{INPUT_EXCEL}'...")
    wb_in = openpyxl.load_workbook(INPUT_EXCEL, read_only=True)
    ws_in = wb_in.active
    
    rows = list(ws_in.iter_rows(values_only=True))
    header = rows[0]
    data = rows[1:]
    print(f"Total input rows: {len(data)}")

    # Step 1: Build dictionary mapping of verified nukta variants
    base_to_nukta_candidates = {}
    corpus_nukta_words = set()

    for r in data:
        w = unicodedata.normalize('NFC', str(r[0]).strip())
        base = strip_nukta(w)
        
        if has_nukta(w):
            corpus_nukta_words.add(w)
            base_to_nukta_candidates.setdefault(base, []).append(w)
            
        for val in [r[6], r[8], r[11], r[14]]: # Rajpal, Hindwi, HinKhoj, Shabdkosh
            if val:
                hw = unicodedata.normalize('NFC', str(val).split()[0].strip()).strip('-:;,()[]')
                if has_nukta(hw) and strip_nukta(hw) == base:
                    base_to_nukta_candidates.setdefault(base, []).append(hw)

    verified_nukta_map = {}
    for base, cands in base_to_nukta_candidates.items():
        in_corpus = [c for c in cands if c in corpus_nukta_words]
        if in_corpus:
            best = max(set(in_corpus), key=in_corpus.count)
        else:
            best = max(set(cands), key=cands.count)
        verified_nukta_map[base] = best

    print(f"Verified dictionary nukta variants mapped: {len(verified_nukta_map)}")

    # Step 2: Create new workbook with Valid_Nukta_Variant column
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Validated_words"

    new_headers = [
        "Word",
        "Valid_Nukta_Variant",
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
    ws_out.append(new_headers)

    dropped_non_nukta_count = 0
    written_count = 0

    for r_idx, r in enumerate(data, start=2):
        w = unicodedata.normalize('NFC', str(r[0]).strip())
        base = strip_nukta(w)
        nukta_variant = verified_nukta_map.get(base, "")

        # Rule: Remove non-nukta variants where the authenticated nukta word already exists in the corpus
        if not has_nukta(w) and nukta_variant and nukta_variant in corpus_nukta_words:
            dropped_non_nukta_count += 1
            continue

        # For kept rows, assemble row with new Valid_Nukta_Variant column at index 1 (Column B)
        # If the word itself has nukta, its valid nukta variant is itself
        val_nukta = w if has_nukta(w) else nukta_variant

        row_values = [
            w,                  # Col A: Word
            val_nukta,          # Col B: Valid_Nukta_Variant (NEW)
            r[1],               # Col C: Overall_Status
            r[2],               # Col D: Mismatch_Type
            r[3],               # Col E: Match_Count
            r[4],               # Col F: Consolidated_Notes
            r[5],               # Col G: Rajpal_Status
            r[6],               # Col H: Rajpal_Entry
            r[7],               # Col I: Hindwi_Status
            r[8],               # Col J: Hindwi_Card_Word
            r[9],               # Col K: Hindwi_Meaning
            r[10],              # Col L: HinKhoj_Status
            r[11],              # Col M: HinKhoj_Card_Word
            r[12],              # Col N: HinKhoj_Meaning
            r[13],              # Col O: Shabdkosh_Status
            r[14],              # Col P: Shabdkosh_H1_Word
            r[15]               # Col Q: Shabdkosh_Meaning
        ]
        ws_out.append(row_values)
        written_count += 1

    print(f"\nSaving new standardized file to '{OUTPUT_EXCEL}'...")
    wb_out.save(OUTPUT_EXCEL)

    print(f"Copying to '{RECHECK_DEST}'...")
    shutil.copy2(OUTPUT_EXCEL, RECHECK_DEST)

    print("\n" + "="*60)
    print("SUCCESSFULLY CREATED STANDARDIZED COPY")
    print(f"Total Words Preserved: {written_count}")
    print(f"Non-nukta Duplicate Rows Removed: {dropped_non_nukta_count}")
    print(f"Saved at: {OUTPUT_EXCEL}")
    print(f"Saved at: {RECHECK_DEST}")
    print("="*60)

if __name__ == '__main__':
    main()
