import os
import sys
import time
import re
from datetime import datetime
import requests
from openpyxl import load_workbook
import config

def log(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted_msg = f"[{timestamp}] {message}"
    print(formatted_msg)
    try:
        with open(config.LOG_FILE, "a", encoding="utf-8") as f:
            f.write(formatted_msg + "\n")
    except Exception as e:
        print(f"Failed to write to log file: {e}")

def normalize_link(link_str):
    if not link_str:
        return ""
    return str(link_str).strip().rstrip("/").lower()

# --- 1. LEETCODE FETCHER ---
def fetch_leetcode_submissions():
    username = getattr(config, "LEETCODE_USERNAME", None)
    if not username:
        return []
    
    url = config.LEETCODE_GRAPHQL_URL
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }
    query_recent = """
    query getRecentAcSubmissions($username: String!, $limit: Int!) {
      recentAcSubmissionList(username: $username, limit: $limit) {
        id
        title
        titleSlug
        timestamp
      }
    }
    """
    payload = {
        "query": query_recent,
        "variables": {"username": username, "limit": 25}
    }
    
    results = []
    try:
        r = requests.post(url, json=payload, headers=headers, timeout=10)
        r.raise_for_status()
        subs = r.json().get("data", {}).get("recentAcSubmissionList", []) or []
        
        query_problem = """
        query selectQuestion($titleSlug: String!) {
          question(titleSlug: $titleSlug) {
            difficulty
            topicTags {
              name
            }
          }
        }
        """
        for sub in subs:
            slug = sub.get("titleSlug")
            raw_title = sub.get("title")
            ts = int(sub.get("timestamp", 0))
            
            diff = "Easy"
            topics = "General"
            try:
                p_r = requests.post(url, json={"query": query_problem, "variables": {"titleSlug": slug}}, headers=headers, timeout=5)
                q = p_r.json().get("data", {}).get("question", {})
                if q:
                    diff = q.get("difficulty", "Easy")
                    tags = [t["name"] for t in q.get("topicTags", []) if "name" in t]
                    if tags:
                        topics = ", ".join(tags)
            except Exception:
                pass
                
            date_str = datetime.fromtimestamp(ts).strftime("%d.%m.%y") if ts else datetime.now().strftime("%d.%m.%y")
            results.append({
                "date": date_str,
                "title": f"{raw_title} - LeetCode",
                "link": f"https://leetcode.com/problems/{slug}/",
                "difficulty": diff,
                "platform": "Leetcode",
                "topics": topics,
                "timestamp": ts
            })
            time.sleep(0.1)
    except Exception as e:
        log(f"Error fetching LeetCode submissions: {e}")
        
    return results

# --- 2. CODEFORCES FETCHER ---
def fetch_codeforces_submissions():
    handle = getattr(config, "CODEFORCES_HANDLE", None)
    if not handle:
        return []
        
    url = f"https://codeforces.com/api/user.status?handle={handle}&from=1&count=50"
    headers = {"User-Agent": "Mozilla/5.0"}
    results = []
    
    try:
        r = requests.get(url, headers=headers, timeout=10)
        r.raise_for_status()
        data = r.json()
        if data.get("status") == "OK":
            submissions = data.get("result", [])
            seen_problems = set()
            
            for sub in submissions:
                if sub.get("verdict") != "OK":
                    continue
                prob = sub.get("problem", {})
                contest_id = prob.get("contestId")
                index = prob.get("index")
                name = prob.get("name")
                
                if not contest_id or not index or not name:
                    continue
                    
                prob_key = f"{contest_id}_{index}"
                if prob_key in seen_problems:
                    continue
                seen_problems.add(prob_key)
                
                rating = prob.get("rating")
                if rating is None or rating <= 1000:
                    diff = "Easy"
                elif rating <= 1600:
                    diff = "Medium"
                else:
                    diff = "Hard"
                    
                tags = prob.get("tags", [])
                topics = ", ".join(tags) if tags else "Competitive Programming"
                
                ts = int(sub.get("creationTimeSeconds", 0))
                date_str = datetime.fromtimestamp(ts).strftime("%d.%m.%y") if ts else datetime.now().strftime("%d.%m.%y")
                
                link = f"https://codeforces.com/contest/{contest_id}/problem/{index}"
                results.append({
                    "date": date_str,
                    "title": f"{name} - Codeforces",
                    "link": link,
                    "difficulty": diff,
                    "platform": "Codeforces",
                    "topics": topics,
                    "timestamp": ts
                })
    except Exception as e:
        log(f"Error fetching Codeforces submissions: {e}")
        
    return results

# --- 3. ATCODER FETCHER ---
def fetch_atcoder_submissions():
    handle = getattr(config, "ATCODER_HANDLE", None)
    if not handle:
        return []
        
    url = f"https://kenkoooo.com/atcoder/atcoder-api/v3/user/submissions?user={handle}&from_second=0"
    headers = {"User-Agent": "Mozilla/5.0"}
    results = []
    
    try:
        r = requests.get(url, headers=headers, timeout=10)
        r.raise_for_status()
        submissions = r.json()
        seen_problems = set()
        
        for sub in submissions:
            if sub.get("result") != "AC":
                continue
            prob_id = sub.get("problem_id")
            contest_id = sub.get("contest_id")
            
            if not prob_id or not contest_id:
                continue
                
            if prob_id in seen_problems:
                continue
            seen_problems.add(prob_id)
            
            parts = prob_id.split("_")
            if len(parts) >= 2:
                task_name = f"{parts[0].upper()} Task {parts[1].upper()}"
                letter = parts[1].lower()
            else:
                task_name = prob_id.upper()
                letter = "a"
                
            if letter in ["a", "b"]:
                diff = "Easy"
            elif letter in ["c", "d"]:
                diff = "Medium"
            else:
                diff = "Hard"
                
            ts = int(sub.get("epoch_second", 0))
            date_str = datetime.fromtimestamp(ts).strftime("%d.%m.%y") if ts else datetime.now().strftime("%d.%m.%y")
            
            link = f"https://atcoder.jp/contests/{contest_id}/tasks/{prob_id}"
            results.append({
                "date": date_str,
                "title": f"{task_name} - AtCoder",
                "link": link,
                "difficulty": diff,
                "platform": "Atcoder",
                "topics": "Competitive Programming",
                "timestamp": ts
            })
    except Exception as e:
        log(f"Error fetching AtCoder submissions: {e}")
        
    return results

# --- 4. GEEKSFORGEEKS FETCHER ---
def fetch_gfg_submissions():
    handle = getattr(config, "GFG_HANDLE", None)
    if not handle:
        return []
        
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    }
    
    urls = [
        f"https://www.geeksforgeeks.org/profile/{handle}?tab=activity",
        f"https://auth.geeksforgeeks.org/user/{handle}/practice/",
        f"https://www.geeksforgeeks.org/user/{handle}/"
    ]
    
    results = []
    seen_links = set()
    
    for url in urls:
        try:
            r = requests.get(url, headers=headers, timeout=10)
            if r.status_code != 200:
                continue
                
            # Parse problem links
            p_matches = re.findall(r'href=[\"\'](https://[^\"]*geeksforgeeks\.org/problems/([^\"]+))[\"\']', r.text)
            for full_link, slug in p_matches:
                clean_link = full_link.split('?')[0].rstrip('/') + '/'
                if clean_link in seen_links:
                    continue
                seen_links.add(clean_link)
                
                title_slug = slug.split('/')[0]
                title = title_slug.replace('-', ' ').title()
                results.append({
                    "date": datetime.now().strftime("%d.%m.%y"),
                    "title": f"{title} - GeeksforGeeks",
                    "link": clean_link,
                    "difficulty": "Medium",
                    "platform": "GeeksforGeeks",
                    "topics": "DSA",
                    "timestamp": int(datetime.now().timestamp())
                })
        except Exception as e:
            log(f"Error fetching GFG submissions from {url}: {e}")
            
    return results

# --- MAIN EXCEL UPDATE PIPELINE ---
def update_excel_tracker():
    log("=== Starting Multi-Platform Tracker Automation ===")
    
    if not os.path.exists(config.EXCEL_FILE):
        log(f"ERROR: Excel file not found at: {config.EXCEL_FILE}")
        return

    all_submissions = []
    
    log("Fetching LeetCode submissions...")
    lc_subs = fetch_leetcode_submissions()
    log(f"-> LeetCode: {len(lc_subs)} AC problem(s) retrieved.")
    all_submissions.extend(lc_subs)
    
    log("Fetching Codeforces submissions...")
    cf_subs = fetch_codeforces_submissions()
    log(f"-> Codeforces: {len(cf_subs)} AC problem(s) retrieved.")
    all_submissions.extend(cf_subs)

    log("Fetching AtCoder submissions...")
    at_subs = fetch_atcoder_submissions()
    log(f"-> AtCoder: {len(at_subs)} AC problem(s) retrieved.")
    all_submissions.extend(at_subs)

    log("Fetching GeeksforGeeks submissions...")
    gfg_subs = fetch_gfg_submissions()
    log(f"-> GeeksforGeeks: {len(gfg_subs)} AC problem(s) retrieved.")
    all_submissions.extend(gfg_subs)

    if not all_submissions:
        log("No submissions found across platforms.")
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

    existing_links = set()
    existing_titles = set()
    
    for r in range(config.START_ROW, sheet.max_row + 1):
        title_val = sheet.cell(row=r, column=2).value
        link_val = sheet.cell(row=r, column=3).value
        if link_val:
            existing_links.add(normalize_link(link_val))
        if title_val:
            existing_titles.add(str(title_val).strip().lower())

    legend_row = None
    for r in range(config.START_ROW, sheet.max_row + 1):
        val = sheet.cell(row=r, column=1).value
        if val and "Legend" in str(val):
            legend_row = r
            break

    target_row = config.START_ROW
    while target_row <= sheet.max_row:
        if legend_row and target_row >= legend_row:
            break
        col_a = sheet.cell(row=target_row, column=1).value
        col_b = sheet.cell(row=target_row, column=2).value
        if col_a is None and col_b is None:
            break
        target_row += 1

    sorted_subs = sorted(all_submissions, key=lambda x: int(x.get("timestamp", 0)))
    
    added_count = 0
    skipped_count = 0

    for item in sorted_subs:
        norm_link = normalize_link(item["link"])
        norm_title = item["title"].strip().lower()
        
        if norm_link in existing_links or norm_title in existing_titles:
            skipped_count += 1
            continue

        if legend_row and target_row >= legend_row:
            sheet.insert_rows(target_row)
            legend_row += 1

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
        
        log(f"-> Added row {target_row} [{item['platform']}]: {item['date']} | {item['title']} | {item['difficulty']} | {item['topics']}")
        added_count += 1
        target_row += 1

    log(f"Summary: {added_count} problem(s) added, {skipped_count} problem(s) already existed.")

    if added_count > 0:
        try:
            workbook.save(config.EXCEL_FILE)
            log(f"SUCCESS: Workbook '{config.EXCEL_FILE}' updated and saved successfully.")
        except PermissionError:
            log(f"ERROR: Permission denied saving '{config.EXCEL_FILE}'. Please close Excel and retry.")
        except Exception as e:
            log(f"ERROR: Failed to save workbook: {e}")
    else:
        log("No changes needed. Workbook is up to date.")

    workbook.close()
    log("=== Execution Finished ===")

if __name__ == "__main__":
    update_excel_tracker()
