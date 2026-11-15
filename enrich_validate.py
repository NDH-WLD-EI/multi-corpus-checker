import os
import sys
import time
import urllib.parse
import re
import html
import csv
import argparse
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import openpyxl

# Force UTF-8 on Windows console
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

INPUT_EXCEL = "corpus_ready_for_enrichment.xlsx"
SAVE_INTERVAL = 25

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9,hi;q=0.8',
    'Referer': 'https://www.google.com/',
}

csv_lock = threading.Lock()

def get_retry_session():
    session = requests.Session()
    retries = Retry(total=3, backoff_factor=1.0, status_forcelist=[429, 500, 502, 503, 504], raise_on_status=False)
    adapter = HTTPAdapter(max_retries=retries, pool_connections=10, pool_maxsize=10)
    session.mount('https://', adapter)
    session.mount('http://', adapter)
    session.headers.update(HEADERS)
    return session

def strip_nukta(text):
    decomposed = unicodedata.normalize('NFD', text)
    stripped = ''.join(c for c in decomposed if c != '\u093c')
    return unicodedata.normalize('NFC', stripped)

def strip_nasal(text):
    return unicodedata.normalize('NFC', text).replace('\u0901', '\u0902')

def check_rekhta(session, word):
    url = f"https://www.rekhtadictionary.com/search?keyword={urllib.parse.quote(word)}&lang=hi"
    try:
        r = session.get(url, timeout=12)
        r.encoding = 'utf-8'
        card_blocks = re.findall(r"<div class=['\"]rdWordCard['\"][^>]*>.*?<h3>(.*?)</h3>.*?<p class=['\"]rdWrdCrdMeaning[^'\"]*['\"]>(.*?)</p>", r.text, re.DOTALL)
        if not card_blocks:
            return "NOT_FOUND", "", ""
            
        norm_query = unicodedata.normalize('NFC', word.strip())
        first_h3 = html.unescape(re.sub(r'<[^>]+>', '', card_blocks[0][0])).strip()
        norm_first_h3 = unicodedata.normalize('NFC', first_h3)
        
        # Strict exact match rule on Box 1
        if norm_query != norm_first_h3:
            return "NOT_FOUND", "", ""
            
        meanings = []
        for h3, m in card_blocks:
            h3_clean = unicodedata.normalize('NFC', html.unescape(re.sub(r'<[^>]+>', '', h3)).strip())
            if h3_clean == norm_query:
                m_clean = html.unescape(re.sub(r'<[^>]+>', '', m)).strip()
                m_clean = re.sub(r'\s+', ' ', m_clean)
                if m_clean and m_clean not in meanings:
                    meanings.append(m_clean)
            else:
                break
                
        if len(meanings) > 1:
            meaning_str = "; ".join(f"[{i+1}] {m}" for i, m in enumerate(meanings))
        elif len(meanings) == 1:
            meaning_str = meanings[0]
        else:
            meaning_str = ""
            
        return "OKAY", norm_first_h3, meaning_str
    except Exception as e:
        return "ERROR", "", str(e)

def check_shabdkosh_meaning(session, word):
    url = f"https://www.shabdkosh.com/search-dictionary?lc=hi&sl=hi&tl=en&e={urllib.parse.quote(word)}"
    try:
        r = session.get(url, timeout=15)
        r.encoding = 'utf-8'
        text = r.text
        
        h1_m = re.search(r'<h1[^>]*id=["\']eil["\'][^>]*>(.*?)</h1>', text, re.DOTALL | re.IGNORECASE)
        if not h1_m:
            h1_m = re.search(r'<h1[^>]*class=["\'][^"\']*pagehead[^"\']*["\'][^>]*>(.*?)</h1>', text, re.DOTALL | re.IGNORECASE)
            
        h1_word = ""
        if h1_m:
            raw_h1 = h1_m.group(1).strip()
            part = re.split(r'<em| - ', raw_h1, flags=re.IGNORECASE)[0]
            h1_word = html.unescape(re.sub(r'<[^>]+>', '', part).strip())
            
        lemmas = re.findall(r'<a[^>]*class=["\'][^"\']*tllemma[^"\']*["\'][^>]*>(.*?)</a>', text, re.IGNORECASE)
        clean_lemmas = []
        for l in lemmas:
            c = html.unescape(re.sub(r'<[^>]+>', '', l)).strip()
            if c and c not in clean_lemmas:
                clean_lemmas.append(c)
                
        norm_w = unicodedata.normalize('NFC', word.strip())
        norm_h1 = unicodedata.normalize('NFC', h1_word.strip()) if h1_word else ""
        
        # Strict validation: must match H1 AND have at least 1 verified lemma definition
        if norm_w == norm_h1 and len(clean_lemmas) > 0:
            return "OKAY", h1_word, ", ".join(clean_lemmas[:3])
        elif len(clean_lemmas) == 0:
            return "NOT_FOUND", "", "N/A"
        elif strip_nukta(norm_w) == strip_nukta(norm_h1):
            return "NUKTA_VARIANT", h1_word, ", ".join(clean_lemmas[:3])
        else:
            return "MISMATCH", h1_word, ", ".join(clean_lemmas[:3])
    except Exception as e:
        return "ERROR", "", str(e)

def load_shard_checkpoint(progress_csv):
    processed = {}
    if os.path.exists(progress_csv):
        try:
            with open(progress_csv, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    w = row.get('Word', '').strip()
                    if w:
                        processed[w] = row
            print(f"[CHECKPOINT] Loaded {len(processed)} existing records from '{progress_csv}'.")
        except Exception as e:
            print(f"[CHECKPOINT WARNING] Error reading checkpoint: {e}")
    return processed

def append_shard_batch(progress_csv, batch, fieldnames):
    with csv_lock:
        file_exists = os.path.exists(progress_csv) and os.path.getsize(progress_csv) > 0
        with open(progress_csv, 'a', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if not file_exists:
                writer.writeheader()
            for rec in batch:
                writer.writerow(rec)

def main():
    parser = argparse.ArgumentParser(description="5-Source Hindi Corpus Enrichment")
    parser.add_argument("--shard-id", type=int, default=1, help="Shard number (1-based)")
    parser.add_argument("--total-shards", type=int, default=5, help="Total number of parallel shards")
    parser.add_argument("--limit", type=int, default=0, help="Limit words (0 = all)")
    parser.add_argument("--workers", type=int, default=3, help="Concurrent threads per shard")
    args = parser.parse_args()

    progress_csv = f"shard_{args.shard_id}_enriched.csv"

    if not os.path.exists(INPUT_EXCEL):
        print(f"Error: '{INPUT_EXCEL}' not found!")
        sys.exit(1)

    print(f"Loading '{INPUT_EXCEL}'...")
    wb = openpyxl.load_workbook(INPUT_EXCEL, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    data = rows[1:]

    total_words = len(data)
    chunk_size = (total_words + args.total_shards - 1) // args.total_shards
    start_idx = (args.shard_id - 1) * chunk_size
    end_idx = min(start_idx + chunk_size, total_words)
    shard_rows = data[start_idx:end_idx]

    if args.limit > 0:
        shard_rows = shard_rows[:args.limit]

    print(f"Shard {args.shard_id}/{args.total_shards}: Assigned {len(shard_rows)} words (Row {start_idx+2} to {end_idx+1})")

    checkpoint = load_shard_checkpoint(progress_csv)
    pending_items = [r for r in shard_rows if unicodedata.normalize('NFC', str(r[0]).strip()) not in checkpoint]

    print(f"Pending in this shard: {len(pending_items)} (already done: {len(checkpoint)})")

    if not pending_items:
        print("All words in this shard are already processed!")
        return

    fieldnames = [
        "Word", "Valid_Nukta_Variant", "Overall_Status", "Mismatch_Type", "Match_Count", "Consolidated_Notes",
        "Rajpal_Status", "Rajpal_Entry",
        "Hindwi_Status", "Hindwi_Card_Word", "Hindwi_Meaning",
        "HinKhoj_Status", "HinKhoj_Card_Word", "HinKhoj_Meaning",
        "Shabdkosh_Status", "Shabdkosh_H1_Word", "Shabdkosh_Meaning",
        "Rekhta_Status", "Rekhta_Word", "Rekhta_Meaning"
    ]

    thread_local = threading.local()
    def get_session():
        if not hasattr(thread_local, "session"):
            thread_local.session = get_retry_session()
        return thread_local.session

    def process_item(r):
        s = get_session()
        w = unicodedata.normalize('NFC', str(r[0]).strip())
        val_nukta = r[1] or ""
        r_stat = r[6] or ""
        r_entry = r[7] or ""
        hw_stat = r[8] or ""
        hw_word = r[9] or ""
        hw_mean = r[10] or ""
        hk_stat = r[11] or ""
        hk_word = r[12] or ""
        hk_mean = r[13] or ""
        
        # 1. Fetch updated Shabdkosh meaning via tllemma
        sk_stat, sk_word, sk_mean = check_shabdkosh_meaning(s, w)
        time.sleep(0.15)
        
        # 2. Fetch Rekhta strictly
        rk_stat, rk_word, rk_mean = check_rekhta(s, w)
        time.sleep(0.15)
        
        # 3. Recalculate 5-source consensus
        all_stats = [r_stat, hw_stat, hk_stat, sk_stat, rk_stat]
        match_count = sum(1 for st in all_stats if st == "OKAY")
        
        if match_count >= 4:
            overall_status = "OKAY"
        elif match_count in (2, 3):
            overall_status = "PARTIAL_MATCH"
        elif match_count == 1:
            overall_status = "MISMATCH"
        else:
            overall_status = "NOT_FOUND"
            
        return {
            "Word": w,
            "Valid_Nukta_Variant": val_nukta,
            "Overall_Status": overall_status,
            "Mismatch_Type": r[3] or "",
            "Match_Count": f"{match_count}/5",
            "Consolidated_Notes": f"Matches {match_count}/5 authoritative sources",
            "Rajpal_Status": r_stat,
            "Rajpal_Entry": r_entry,
            "Hindwi_Status": hw_stat,
            "Hindwi_Card_Word": hw_word,
            "Hindwi_Meaning": hw_mean,
            "HinKhoj_Status": hk_stat,
            "HinKhoj_Card_Word": hk_word,
            "HinKhoj_Meaning": hk_mean,
            "Shabdkosh_Status": sk_stat,
            "Shabdkosh_H1_Word": sk_word,
            "Shabdkosh_Meaning": sk_mean,
            "Rekhta_Status": rk_stat,
            "Rekhta_Word": rk_word,
            "Rekhta_Meaning": rk_mean
        }

    batch = []
    completed = len(checkpoint)
    t_start = time.time()

    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {executor.submit(process_item, r): r for r in pending_items}
        for future in as_completed(futures):
            res = future.result()
            w = res['Word']
            checkpoint[w] = res
            batch.append(res)
            completed += 1

            if completed % 10 == 0 or completed == len(shard_rows):
                elapsed = time.time() - t_start
                rate = (completed - len(checkpoint) + len(batch)) / elapsed if elapsed > 0 else 0
                print(f"[Shard {args.shard_id}] [{completed}/{len(shard_rows)}] '{w:12s}' -> {res['Overall_Status']:14s} ({res['Match_Count']}) | Speed: {rate:.1f} w/s")

            if len(batch) >= SAVE_INTERVAL:
                append_shard_batch(progress_csv, batch, fieldnames)
                batch.clear()

    if batch:
        append_shard_batch(progress_csv, batch, fieldnames)
        batch.clear()

    print(f"\n[Shard {args.shard_id} COMPLETE] All {completed} words verified and saved to '{progress_csv}'!")

if __name__ == '__main__':
    main()
