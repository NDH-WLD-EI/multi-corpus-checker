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

# Force UTF-8 on Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

INPUT_EXCEL = "Validated_content_added.xlsx"
FTS_FILE = "fts.txt"
SHEET_NAME = "Validated_words"
SAVE_INTERVAL = 25

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9,hi;q=0.8',
    'Referer': 'https://www.google.com/',
}

csv_lock = threading.Lock()


def get_retry_session():
    """Create requests session with robust retry logic."""
    session = requests.Session()
    retries = Retry(
        total=3,
        backoff_factor=1.0,
        status_forcelist=[429, 500, 502, 503, 504],
        raise_on_status=False
    )
    adapter = HTTPAdapter(max_retries=retries, pool_connections=10, pool_maxsize=10)
    session.mount('https://', adapter)
    session.mount('http://', adapter)
    session.headers.update(HEADERS)
    return session


def load_rajpal_index(filepath):
    """Loads Rajpal Hindi Shabdakosh OCR text and indexes all headwords in memory."""
    print(f"[RAJPAL] Loading and indexing '{filepath}'...")
    t0 = time.time()
    exact_index = {}
    lines_list = []
    
    if not os.path.exists(filepath):
        print(f"[RAJPAL WARNING] '{filepath}' not found! Rajpal checks will be skipped.")
        return exact_index, lines_list

    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        for idx, line in enumerate(f, 1):
            s = line.strip()
            if not s:
                continue
            lines_list.append((idx, s))
            m = re.match(r'^([^\s\-—–\[\(:]+)[\s\-—–\[\(:]', s)
            if m:
                w_norm = unicodedata.normalize('NFC', m.group(1).strip())
                if w_norm not in exact_index:
                    exact_index[w_norm] = s

    t1 = time.time()
    print(f"[RAJPAL] Indexed {len(exact_index)} headwords in {t1 - t0:.2f}s.")
    return exact_index, lines_list


def check_rajpal(word, rajpal_index):
    """Check word against Rajpal index."""
    norm_w = unicodedata.normalize('NFC', word.strip())
    entry = rajpal_index.get(norm_w)
    
    if entry:
        meaning = entry.split('-', 1)[1].strip() if '-' in entry else entry
        return "OKAY", norm_w, meaning
        
    # Check alternate spelling (e.g. anusvara vs chandrabindu)
    alt_cand = norm_w.replace('\u0901', '\u0902') if '\u0901' in norm_w else norm_w.replace('\u0902', '\u0901')
    entry_alt = rajpal_index.get(alt_cand)
    if entry_alt:
        m_alt = entry_alt.split('-', 1)[1].strip() if '-' in entry_alt else entry_alt
        return "MISMATCH", alt_cand, m_alt
        
    return "NOT_FOUND", "", ""


def check_hindwi(session, word):
    """Check Hindwi Dictionary: first card word & meaning."""
    url = f"https://www.hindwidictionary.com/search?keyword={urllib.parse.quote(word)}"
    try:
        r = session.get(url, timeout=12)
        resp_html = r.text
        card_match = re.search(
            r'<div class=["\']WordRelatedLists saveWord listingCardBg overflowAuto["\']>(.*?)</div>',
            resp_html,
            re.DOTALL
        )
        if card_match:
            card_html = card_match.group(1)
            li_match = re.search(
                r'<li>\s*([^<\n\r]+?)\s*<p class=["\']nobullet["\']>(.*?)</p>',
                card_html,
                re.DOTALL
            )
            if li_match:
                card_word = html.unescape(li_match.group(1).strip())
                raw_m = li_match.group(2).strip()
                meaning = html.unescape(re.sub(r'\s+', ' ', raw_m))
                
                norm_w = unicodedata.normalize('NFC', word.strip())
                norm_card = unicodedata.normalize('NFC', card_word.strip())
                status = "OKAY" if norm_w == norm_card else "MISMATCH"
                return status, card_word, meaning
        return "NOT_FOUND", "", ""
    except Exception as e:
        return "ERROR", "", str(e)


def check_hinkhoj(session, word):
    """Check HinKhoj Dictionary: first card word & meaning."""
    url = f"https://dict.hinkhoj.com/{urllib.parse.quote(word)}-meaning-in-english.words"
    try:
        r = session.get(url, timeout=12)
        r.encoding = 'utf-8'
        resp_html = r.text
        card_match = re.search(r'<div[^>]*class=[\'"][^\'"]*word_mean_desc_css[^\'"]*[\'"][^>]*>(.*?)</div>', resp_html, re.DOTALL | re.IGNORECASE)
        if card_match:
            card_html = card_match.group(1)
            equal_sep_match = re.search(r'^(.*?)\s*<span[^>]*class=[\'"][^\'"]*equal_sep[^\'"]*[\'"][^>]*>', card_html.strip(), re.DOTALL | re.IGNORECASE)
            if equal_sep_match:
                card_word = re.sub(r'<[^>]+>', '', equal_sep_match.group(1)).strip()
            else:
                card_word = card_html.split('=')[0].strip()
                card_word = re.sub(r'<[^>]+>', '', card_word).strip()
                
            card_word = html.unescape(card_word)
            meanings = re.findall(r'<a[^>]*class=[\'"][^\'"]*eng_dict_span[^\'"]*[\'"][^>]*>(.*?)</a>', card_html, re.DOTALL | re.IGNORECASE)
            cleaned_meanings = [html.unescape(re.sub(r'<[^>]+>', '', m).strip()) for m in meanings]
            meaning_str = ", ".join(cleaned_meanings) if cleaned_meanings else "N/A"
            
            norm_w = unicodedata.normalize('NFC', word.strip())
            norm_card = unicodedata.normalize('NFC', card_word.strip())
            status = "OKAY" if norm_w == norm_card else "MISMATCH"
            return status, card_word, meaning_str
        return "NOT_FOUND", "", ""
    except Exception as e:
        return "ERROR", "", str(e)


def check_shabdkosh(session, word):
    """Check Shabdkosh: H1 header word & top meaning."""
    url = f"https://www.shabdkosh.com/search-dictionary?lc=hi&sl=en&tl=hi&e={urllib.parse.quote(word)}"
    try:
        r = session.get(url, timeout=15)
        r.encoding = 'utf-8'
        resp_html = r.text
        h1_match = re.search(r'<h1[^>]*id=["\']eil["\'][^>]*>(.*?)</h1>', resp_html, re.DOTALL | re.IGNORECASE)
        if not h1_match:
            h1_match = re.search(r'<h1[^>]*class=["\'][^"\']*pagehead[^"\']*["\'][^>]*>(.*?)</h1>', resp_html, re.DOTALL | re.IGNORECASE)
            
        if h1_match:
            h1_raw = h1_match.group(1).strip()
            part = re.split(r'<em| - ', h1_raw, flags=re.IGNORECASE)[0]
            h1_word = html.unescape(re.sub(r'<[^>]+>', '', part).strip())
            
            defs = re.findall(r'<a[^>]*class=["\'][^"\']*inlinelinks[^"\']*["\'][^>]*>(.*?)</a>', resp_html, re.IGNORECASE)
            defs_clean = [re.sub(r'<[^>]+>', '', d).strip() for d in defs if d.strip()][:3]
            meaning_str = ", ".join(defs_clean) if defs_clean else "N/A"
            
            norm_w = unicodedata.normalize('NFC', word.strip())
            norm_h1 = unicodedata.normalize('NFC', h1_word.strip())
            status = "OKAY" if norm_w == norm_h1 else "MISMATCH"
            return status, h1_word, meaning_str
        return "NOT_FOUND", "", ""
    except Exception as e:
        return "ERROR", "", str(e)


def validate_word(session, rajpal_index, row_num, word):
    """Validates word across all 4 sources and determines consolidated status."""
    # 1. Rajpal (instant local lookup)
    r_stat, r_word, r_mean = check_rajpal(word, rajpal_index)
    
    # 2. Hindwi
    hw_stat, hw_word, hw_mean = check_hindwi(session, word)
    
    # 3. HinKhoj
    hk_stat, hk_word, hk_mean = check_hinkhoj(session, word)
    
    # 4. Shabdkosh
    sk_stat, sk_word, sk_mean = check_shabdkosh(session, word)
    
    # Check match count
    match_count = sum(1 for s in [r_stat, hw_stat, hk_stat, sk_stat] if s == "OKAY")
    
    if match_count == 4:
        overall_status = "OKAY"
        consolidated_note = "OKAY: Confirmed in all 4 sources"
    elif match_count >= 2:
        overall_status = "PARTIAL_MATCH"
        notes = []
        if r_stat != "OKAY" and r_word: notes.append(f"Rajpal: {r_word}")
        if hw_stat != "OKAY" and hw_word: notes.append(f"Hindwi: {hw_word}")
        if hk_stat != "OKAY" and hk_word: notes.append(f"HinKhoj: {hk_word}")
        if sk_stat != "OKAY" and sk_word: notes.append(f"Shabdkosh: {sk_word}")
        consolidated_note = f"Matches {match_count}/4. Differs in: " + (", ".join(notes) if notes else "N/A")
    else:
        overall_status = "MISMATCH"
        consolidated_note = f"Matches {match_count}/4. Low consensus across sources"
        
    return {
        'Index': row_num,
        'Word': word,
        'Overall_Status': overall_status,
        'Match_Count': f"{match_count}/4",
        'Consolidated_Notes': consolidated_note,
        'Rajpal_Status': r_stat,
        'Rajpal_Entry': r_mean[:200] if r_mean else "",
        'Hindwi_Status': hw_stat,
        'Hindwi_Card_Word': hw_word,
        'Hindwi_Meaning': hw_mean[:200] if hw_mean else "",
        'HinKhoj_Status': hk_stat,
        'HinKhoj_Card_Word': hk_word,
        'HinKhoj_Meaning': hk_mean[:200] if hk_mean else "",
        'Shabdkosh_Status': sk_stat,
        'Shabdkosh_H1_Word': sk_word,
        'Shabdkosh_Meaning': sk_mean[:200] if sk_mean else ""
    }


def load_shard_checkpoint(progress_csv):
    """Load existing checkpoint to resume without repeating work."""
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


def append_shard_batch(progress_csv, batch):
    """Appends records to the shard CSV file safely."""
    with csv_lock:
        file_exists = os.path.exists(progress_csv) and os.path.getsize(progress_csv) > 0
        fieldnames = [
            'Index', 'Word', 'Overall_Status', 'Match_Count', 'Consolidated_Notes',
            'Rajpal_Status', 'Rajpal_Entry',
            'Hindwi_Status', 'Hindwi_Card_Word', 'Hindwi_Meaning',
            'HinKhoj_Status', 'HinKhoj_Card_Word', 'HinKhoj_Meaning',
            'Shabdkosh_Status', 'Shabdkosh_H1_Word', 'Shabdkosh_Meaning'
        ]
        with open(progress_csv, 'a', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if not file_exists:
                writer.writeheader()
            for rec in batch:
                writer.writerow(rec)


def main():
    parser = argparse.ArgumentParser(description="Multi-Source Hindi Corpus Validator")
    parser.add_argument("--shard-id", type=int, default=1, help="Shard number (1-based)")
    parser.add_argument("--total-shards", type=int, default=1, help="Total number of parallel shards")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of words to check (0 = all in shard)")
    parser.add_argument("--workers", type=int, default=3, help="Concurrent threads per shard runner")
    args = parser.parse_args()

    progress_csv = f"shard_{args.shard_id}_results.csv"

    # 1. Load Excel words
    if not os.path.exists(INPUT_EXCEL):
        print(f"Error: '{INPUT_EXCEL}' not found!")
        sys.exit(1)

    print(f"Loading '{INPUT_EXCEL}'...")
    wb = openpyxl.load_workbook(INPUT_EXCEL, data_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    
    all_words = []
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if r_idx == 1:
            continue  # header
        if not row or not row[0]:
            continue
        w = str(row[0]).strip()
        all_words.append((r_idx, w))

    total_words = len(all_words)
    print(f"Total words in corpus: {total_words}")

    # 2. Slice shard
    chunk_size = (total_words + args.total_shards - 1) // args.total_shards
    start_idx = (args.shard_id - 1) * chunk_size
    end_idx = min(start_idx + chunk_size, total_words)
    shard_words = all_words[start_idx:end_idx]

    if args.limit > 0:
        shard_words = shard_words[:args.limit]

    print(f"Shard {args.shard_id}/{args.total_shards}: Assigned {len(shard_words)} words (Row {start_idx+2} to {end_idx+1})")

    # 3. Load Rajpal Index once
    rajpal_index, _ = load_rajpal_index(FTS_FILE)

    # 4. Load Checkpoint
    checkpoint = load_shard_checkpoint(progress_csv)
    pending_items = [item for item in shard_words if item[1] not in checkpoint]

    print(f"Pending to process in this shard: {len(pending_items)} (already done: {len(checkpoint)})")

    if not pending_items:
        print("All words in this shard are already processed!")
        return

    # 5. Process concurrently
    thread_local = threading.local()
    def get_session():
        if not hasattr(thread_local, "session"):
            thread_local.session = get_retry_session()
        return thread_local.session

    def process_item(item):
        s = get_session()
        r_num, w = item
        res = validate_word(s, rajpal_index, r_num, w)
        time.sleep(0.2)  # polite pause
        return res

    batch = []
    completed_in_shard = len(checkpoint)
    t_start = time.time()

    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {executor.submit(process_item, item): item for item in pending_items}
        
        for future in as_completed(futures):
            res = future.result()
            w = res['Word']
            checkpoint[w] = res
            batch.append(res)
            completed_in_shard += 1

            if completed_in_shard % 10 == 0 or completed_in_shard == len(shard_words):
                elapsed = time.time() - t_start
                rate = (completed_in_shard - len(checkpoint) + len(batch)) / elapsed if elapsed > 0 else 0
                print(f"[Shard {args.shard_id}] [{completed_in_shard}/{len(shard_words)}] '{w:12s}' -> {res['Overall_Status']:14s} ({res['Match_Count']}) | Speed: {rate:.1f} w/s")

            if len(batch) >= SAVE_INTERVAL:
                append_shard_batch(progress_csv, batch)
                batch.clear()

    if batch:
        append_shard_batch(progress_csv, batch)
        batch.clear()

    print(f"\n[Shard {args.shard_id} COMPLETE] All {completed_in_shard} words verified and saved to '{progress_csv}'!")


if __name__ == '__main__':
    main()
