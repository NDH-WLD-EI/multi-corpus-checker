import os
import sys
import time
import urllib.parse
import re
import html
import csv
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import openpyxl

INPUT_EXCEL = "corpus_ready_for_enrichment.xlsx"
OUTPUT_CSV = "shard_remaining_enriched.csv"
SAVE_INTERVAL = 10

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9,hi;q=0.8',
    'Referer': 'https://www.google.com/',
}

csv_lock = threading.Lock()

def get_retry_session():
    session = requests.Session()
    retries = Retry(total=3, backoff_factor=0.8, status_forcelist=[429, 500, 502, 503, 504], raise_on_status=False)
    adapter = HTTPAdapter(max_retries=retries, pool_connections=12, pool_maxsize=12)
    session.mount('https://', adapter)
    session.mount('http://', adapter)
    session.headers.update(HEADERS)
    return session

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
        r = session.get(url, timeout=12)
        r.encoding = 'utf-8'
        text = r.text
        
        h1_m = re.search(r'<h1[^>]*id=["\']eil["\'][^>]*>(.*?)</h1>', text, re.DOTALL | re.IGNORECASE)
        if not h1_m:
            return "NOT_FOUND", "", "N/A"
            
        h1_text = html.unescape(re.sub(r'<[^>]+>', '', h1_m.group(1))).strip()
        h1_word = h1_text.split(' - ')[0].strip() if ' - ' in h1_text else h1_text
        norm_h1_word = unicodedata.normalize('NFC', h1_word)
        norm_query = unicodedata.normalize('NFC', word.strip())
        
        raw_lemmas = re.findall(r'<a[^>]*class=["\'][^"\']*tllemma[^"\']*["\'][^>]*>(.*?)</a>', text, re.DOTALL)
        lemmas = []
        for l in raw_lemmas:
            clean = html.unescape(re.sub(r'<[^>]+>', '', l)).strip()
            if clean and clean not in lemmas:
                lemmas.append(clean)
                
        if norm_h1_word == norm_query and len(lemmas) > 0:
            meaning_str = ", ".join(lemmas[:5])
            return "OKAY", norm_h1_word, meaning_str
        elif len(lemmas) > 0:
            return "NOT_FOUND", norm_h1_word, "N/A"
        else:
            return "NOT_FOUND", norm_h1_word, "N/A"
    except Exception as e:
        return "ERROR", "", str(e)

def load_checkpoint(filepath):
    checkpoint = {}
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                w = unicodedata.normalize('NFC', row.get('Word', '').strip())
                if w:
                    checkpoint[w] = row
    return checkpoint

def append_batch(filepath, batch, fieldnames):
    with csv_lock:
        file_exists = os.path.exists(filepath) and os.path.getsize(filepath) > 0
        with open(filepath, 'a', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if not file_exists:
                writer.writeheader()
            for rec in batch:
                writer.writerow(rec)

def main():
    import glob
    print(f"Loading '{INPUT_EXCEL}'...")
    wb = openpyxl.load_workbook(INPUT_EXCEL, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]

    # Collect already done words from all existing shards
    done = set()
    for f in glob.glob("shard_*_enriched.csv"):
        if f == OUTPUT_CSV:
            continue
        with open(f, 'r', encoding='utf-8') as fp:
            for r in csv.DictReader(fp):
                w = unicodedata.normalize('NFC', r.get('Word', '').strip())
                if w:
                    done.add(w)

    checkpoint = load_checkpoint(OUTPUT_CSV)
    for w in checkpoint:
        done.add(w)

    pending_items = [r for r in data if unicodedata.normalize('NFC', str(r[0]).strip()) not in done]
    print(f"Total base words: {len(data)}")
    print(f"Already done across shards: {len(done)}")
    print(f"Remaining to process: {len(pending_items)}")

    if not pending_items:
        print("[SUCCESS] All words are already processed!")
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
        
        sk_stat, sk_word, sk_mean = check_shabdkosh_meaning(s, w)
        time.sleep(0.08)
        rk_stat, rk_word, rk_mean = check_rekhta(s, w)
        time.sleep(0.08)
        
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
    completed = 0
    total = len(pending_items)
    t_start = time.time()

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(process_item, r): r for r in pending_items}
        for future in as_completed(futures):
            res = future.result()
            w = res['Word']
            batch.append(res)
            completed += 1

            if completed % 10 == 0 or completed == total:
                elapsed = time.time() - t_start
                rate = completed / elapsed if elapsed > 0 else 0
                eta = (total - completed) / rate if rate > 0 else 0
                print(f"[{completed}/{total}] '{w:12s}' -> {res['Overall_Status']:14s} ({res['Match_Count']}) | Speed: {rate:.1f} w/s | ETA: {eta:.0f}s")

            if len(batch) >= SAVE_INTERVAL:
                append_batch(OUTPUT_CSV, batch, fieldnames)
                batch.clear()

    if batch:
        append_batch(OUTPUT_CSV, batch, fieldnames)
        batch.clear()

    print(f"\n[COMPLETE] Successfully processed all {total} remaining words and saved to '{OUTPUT_CSV}'!")

if __name__ == '__main__':
    main()
