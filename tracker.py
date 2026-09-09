import os, sys, time, re
from datetime import datetime
import requests
import openpyxl
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
import config

def log(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted_msg = f"[{timestamp}] {message}"
    print(formatted_msg)
    sys.stdout.flush()
    try:
        with open(config.LOG_FILE, "a", encoding="utf-8") as f:
            f.write(formatted_msg + "\n")
    except Exception as e:
        print(f"Failed to write to log file: {e}")

def normalize_link(link_str):
    if not link_str:
        return ""
    return str(link_str).strip().rstrip("/").lower()

def fetch_leetcode_submissions(existing_links, existing_titles):
    username = getattr(config, "LEETCODE_USERNAME", None)
    if not username:
        return []
    url = config.LEETCODE_GRAPHQL_URL
    headers = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
    query_recent = """
    query getRecentAcSubmissions($username: String!, $limit: Int!) {
      recentAcSubmissionList(username: $username, limit: $limit) {
        id title titleSlug timestamp
      }
    }
    """
    query_problem = """
    query selectQuestion($titleSlug: String!) {
      question(titleSlug: $titleSlug) {
        difficulty topicTags { name }
      }
    }
    """
    results = []
    try:
        r = requests.post(url, json={"query": query_recent, "variables": {"username": username, "limit": 25}}, headers=headers, timeout=10)
        r.raise_for_status()
        subs = r.json().get("data", {}).get("recentAcSubmissionList", []) or []
        
        for sub in subs:
            slug = sub.get("titleSlug")
            raw_title = sub.get("title")
            ts = int(sub.get("timestamp", 0))
            if not slug or not raw_title:
                continue
            
            link = f"https://leetcode.com/problems/{slug}/"
            full_title = f"{raw_title} - LeetCode"
            
            if normalize_link(link) in existing_links or full_title.strip().lower() in existing_titles:
                continue

            diff, topics = "Medium", "General"
            try:
                p_r = requests.post(url, json={"query": query_problem, "variables": {"titleSlug": slug}}, headers=headers, timeout=5)
                q = p_r.json().get("data", {}).get("question", {})
                if q:
                    diff = q.get("difficulty", "Medium")
                    tags = [t["name"] for t in q.get("topicTags", []) if "name" in t]
                    if tags:
                        topics = ", ".join(tags)
            except Exception as e:
                log(f"Warning: Could not fetch details for LC {slug}: {e}")

            date_str = datetime.fromtimestamp(ts).strftime("%d.%m.%y") if ts else datetime.now().strftime("%d.%m.%y")
            results.append({
                "date": date_str,
                "title": full_title,
                "link": link,
                "difficulty": diff,
                "platform": "Leetcode",
                "topics": topics,
                "timestamp": ts
            })
    except Exception as e:
        log(f"Error fetching LeetCode submissions: {e}")
    return results

def fetch_codeforces_submissions(existing_links, existing_titles):
    handle = getattr(config, "CODEFORCES_HANDLE", None)
    if not handle:
        return []
    results = []
    try:
        r = requests.get(f"https://codeforces.com/api/user.status?handle={handle}&from=1&count=50", headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        r.raise_for_status()
        data = r.json()
        if data.get("status") == "OK":
            seen = set()
            for sub in data.get("result", []):
                if sub.get("verdict") != "OK":
                    continue
                prob = sub.get("problem", {})
                contest_id = prob.get("contestId")
                index = prob.get("index")
                name = prob.get("name")
                if not all([contest_id, index, name]):
                    continue
                key = f"{contest_id}_{index}"
                if key in seen:
                    continue
                seen.add(key)
                
                link = f"https://codeforces.com/contest/{contest_id}/problem/{index}"
                full_title = f"{name} - Codeforces"
                
                if normalize_link(link) in existing_links or full_title.strip().lower() in existing_titles:
                    continue
                
                rating = prob.get("rating")
                diff = "Easy" if rating is None or rating <= 1000 else ("Medium" if rating <= 1600 else "Hard")
                tags = prob.get("tags", [])
                topics = ", ".join(tags) if tags else "Competitive Programming"
                ts = int(sub.get("creationTimeSeconds", 0))
                date_str = datetime.fromtimestamp(ts).strftime("%d.%m.%y") if ts else datetime.now().strftime("%d.%m.%y")
                results.append({
                    "date": date_str,
                    "title": full_title,
                    "link": link,
                    "difficulty": diff,
                    "platform": "Codeforces",
                    "topics": topics,
                    "timestamp": ts
                })
    except Exception as e:
        log(f"Error fetching Codeforces submissions: {e}")
    return results

def fetch_atcoder_submissions(existing_links, existing_titles):
    handle = getattr(config, "ATCODER_HANDLE", None)
    if not handle:
        return []
    results = []
    try:
        r = requests.get(f"https://kenkoooo.com/atcoder/atcoder-api/v3/user/submissions?user={handle}&from_second=0", headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        r.raise_for_status()
        seen = set()
        for sub in r.json():
            if sub.get("result") != "AC":
                continue
            prob_id = sub.get("problem_id")
            contest_id = sub.get("contest_id")
            if not prob_id or not contest_id or prob_id in seen:
                continue
            seen.add(prob_id)
            
            link = f"https://atcoder.jp/contests/{contest_id}/tasks/{prob_id}"
            parts = prob_id.split("_")
            task_name = f"{parts[0].upper()} Task {parts[1].upper()}" if len(parts) >= 2 else prob_id.upper()
            full_title = f"{task_name} - AtCoder"
            
            if normalize_link(link) in existing_links or full_title.strip().lower() in existing_titles:
                continue
            
            letter = parts[1].lower() if len(parts) >= 2 else "a"
            diff = "Easy" if letter in ["a", "b"] else ("Medium" if letter in ["c", "d"] else "Hard")
            ts = int(sub.get("epoch_second", 0))
            date_str = datetime.fromtimestamp(ts).strftime("%d.%m.%y") if ts else datetime.now().strftime("%d.%m.%y")
            results.append({
                "date": date_str,
                "title": full_title,
                "link": link,
                "difficulty": diff,
                "platform": "Atcoder",
                "topics": "Competitive Programming",
                "timestamp": ts
            })
    except Exception as e:
        log(f"Error fetching AtCoder submissions: {e}")
    return results

def fetch_gfg_submissions(existing_links, existing_titles):
    return []

def unmerge_empty_ranges(sheet):
    to_unmerge = []
    for rng in list(sheet.merged_cells.ranges):
        if rng.min_row >= config.START_ROW:
            to_unmerge.append(rng)
    for rng in to_unmerge:
        try:
            sheet.unmerge_cells(str(rng))
        except Exception:
            pass

def update_excel_tracker():
    lock_file = os.path.join(os.path.dirname(config.LOG_FILE), "tracker.lock")
    if os.path.exists(lock_file):
        try:
            mtime = os.path.getmtime(lock_file)
            if time.time() - mtime > 300:
                os.remove(lock_file)
            else:
                with open(lock_file, "r", encoding="utf-8") as f:
                    pid = f.read().strip()
                log(f"Another instance is running (PID {pid}). Exiting.")
                return
        except Exception:
            pass

    try:
        with open(lock_file, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
        _run()
    finally:
        if os.path.exists(lock_file):
            try:
                os.remove(lock_file)
            except Exception:
                pass

def _run():
    log("=== Starting Multi-Platform Tracker Automation ===")

    if not os.path.exists(config.EXCEL_FILE):
        log(f"ERROR: Excel file not found at: {config.EXCEL_FILE}")
        return

    try:
        workbook = load_workbook(config.EXCEL_FILE)
    except PermissionError:
        log(f"ERROR: Permission denied accessing '{config.EXCEL_FILE}'. Please close Excel if open.")
        return
    except Exception as e:
        log(f"ERROR: Failed to open Excel file: {e}")
        return

    if config.SHEET_NAME not in workbook.sheetnames:
        log(f"ERROR: Sheet '{config.SHEET_NAME}' not found.")
        workbook.close()
        return

    sheet = workbook[config.SHEET_NAME]

    # Clean up any merged ranges in data area
    unmerge_empty_ranges(sheet)

    existing_links = set()
    existing_titles = set()
    last_data_row = config.START_ROW - 1

    scan_max = max(sheet.max_row + 50, 100)
    for r in range(config.START_ROW, scan_max):
        title_val = sheet.cell(row=r, column=2).value
        link_val = sheet.cell(row=r, column=3).value
        if link_val:
            existing_links.add(normalize_link(link_val))
        if title_val:
            t_str = str(title_val).strip()
            existing_titles.add(t_str.lower())
            if not any(k in t_str.lower() for k in ["legend", "yellow", "enter", "total", "count"]):
                last_data_row = max(last_data_row, r)

    log(f"Pre-scan complete: {len(existing_links)} existing links found. Last data row: {last_data_row}")

    all_submissions = []
    
    log("Fetching LeetCode submissions...")
    lc = fetch_leetcode_submissions(existing_links, existing_titles)
    log(f"-> LeetCode: {len(lc)} NEW submission(s) to add.")
    all_submissions.extend(lc)

    log("Fetching Codeforces submissions...")
    cf = fetch_codeforces_submissions(existing_links, existing_titles)
    log(f"-> Codeforces: {len(cf)} NEW submission(s) to add.")
    all_submissions.extend(cf)

    log("Fetching AtCoder submissions...")
    at = fetch_atcoder_submissions(existing_links, existing_titles)
    log(f"-> AtCoder: {len(at)} NEW submission(s) to add.")
    all_submissions.extend(at)

    log("Fetching GeeksforGeeks submissions...")
    gfg = fetch_gfg_submissions(existing_links, existing_titles)
    log(f"-> GeeksforGeeks: {len(gfg)} NEW submission(s) to add.")
    all_submissions.extend(gfg)

    if not all_submissions:
        log("No new submissions to insert. Workbook is up to date.")
        workbook.close()
        log("=== Execution Finished ===")
        return

    sorted_subs = sorted(all_submissions, key=lambda x: int(x.get("timestamp", 0)))
    target_row = last_data_row + 1
    added_count = 0

    for item in sorted_subs:
        norm_link = normalize_link(item["link"])
        norm_title = item["title"].strip().lower()

        if norm_link in existing_links or norm_title in existing_titles:
            continue

        sheet.cell(row=target_row, column=1).value = item["date"]
        sheet.cell(row=target_row, column=2).value = item["title"]
        sheet.cell(row=target_row, column=3).value = item["link"]
        sheet.cell(row=target_row, column=4).value = item["difficulty"]
        sheet.cell(row=target_row, column=5).value = item["platform"]
        sheet.cell(row=target_row, column=6).value = item["topics"]
        sheet.cell(row=target_row, column=7).value = 1
        sheet.cell(row=target_row, column=8).value = "=SUM($G$14:$G$1000)"

        existing_links.add(norm_link)
        existing_titles.add(norm_title)

        log(f"-> Added row {target_row} [{item['platform']}]: {item['date']} | {item['title']} | {item['difficulty']}")
        added_count += 1
        target_row += 1

    log(f"Summary: {added_count} new problem(s) added.")

    if added_count > 0:
        try:
            workbook.save(config.EXCEL_FILE)
            log(f"SUCCESS: Saved '{config.EXCEL_FILE}'.")
        except PermissionError:
            log(f"ERROR: Permission denied saving '{config.EXCEL_FILE}'. Please close Excel and re-run.")
        except Exception as e:
            log(f"ERROR saving: {e}")

    workbook.close()
    log("=== Execution Finished ===")

if __name__ == "__main__":
    update_excel_tracker()
