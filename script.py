import os
import re
import json
import time
import random
import urllib.parse
import smtplib
from email.mime.text import MIMEText

print("✅ Standard library imports successful", flush=True)

try:
    import feedparser
    print("✅ feedparser imported", flush=True)
except ImportError as e:
    print(f"❌ Failed to import feedparser: {e}", flush=True)
    raise

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    print("✅ openpyxl imported", flush=True)
except ImportError as e:
    print(f"❌ Failed to import openpyxl: {e}", flush=True)
    raise

try:
    import requests
    print("✅ requests imported", flush=True)
except ImportError as e:
    print(f"❌ Failed to import requests: {e}", flush=True)
    raise

try:
    from google import genai
    print("✅ google-genai imported", flush=True)
except ImportError as e:
    print(f"❌ Failed to import google-genai: {e}", flush=True)
    raise

print("✅ All imports successful", flush=True)

# ==========================================
# CONFIGURATION & ENVIRONMENT SECRETS
# ==========================================
BASE_DIR               = os.path.dirname(os.path.abspath(__file__))
INPUT_KEYWORDS_FILE    = os.path.join(BASE_DIR, "keywords.txt")
INPUT_COMPANIES_FILE   = os.path.join(BASE_DIR, "companies.txt")   # ← NEW
OUTPUT_EXCEL_FILE      = os.path.join(BASE_DIR, "docs", "data", f"rxbenefits_intel_hub_report_{time.strftime('%Y-%m-%d')}.xlsx")
OUTPUT_JSON_FILE       = os.path.join(BASE_DIR, "docs", "data", f"rxbenefits_{time.strftime('%Y-%m-%d')}.json")
GEMINI_API_KEY         = os.environ.get("GEMINI_API_KEY", "").strip()
WP_SITE_URL            = os.environ.get("WP_SITE_URL", "").strip()
WP_USERNAME            = os.environ.get("WP_USERNAME", "").strip()
WP_APP_PASS            = os.environ.get("WP_APP_PASS", "").strip()
SENDER_EMAIL           = os.environ.get("SENDER_EMAIL", "").strip()
SENDER_APP_PASSWORD    = os.environ.get("SENDER_APP_PASSWORD", "").strip()
RECIPIENT_EMAIL        = os.environ.get("RECIPIENT_EMAIL", "").strip()

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36"
}

# ==========================================
# AI PROMPTS  (two-stage filter)
# ==========================================

# Stage 1 — Batch scoring: cheap, one call per 15 articles
SCORING_PROMPT = """
You are a relevance analyst for RxBenefits, a pharmacy benefits management (PBM)
company that serves self-funded employers.

You will receive a JSON list of news article headlines and short snippets.
Score each one from 1 to 5 for relevance to RxBenefits' business:

5 = Must-include: Directly about a named PBM (CVS Caremark, OptumRx, Express Scripts,
    Capital Rx, etc.), PBM reform or legislation, drug pricing policy, employer pharmacy
    benefit strategy, specialty drug coverage, or biosimilars affecting plan sponsors.
4 = High value: General PBM industry news, drug pricing trends that affect employers,
    formulary or prior-authorization policy changes, self-funded employer benefit strategy.
3 = Moderate: Tangentially related healthcare cost news that could affect employer plans.
2 = Low: General pharma or drug news with no direct PBM or employer-benefit angle.
1 = Irrelevant: Clinical drug research, consumer health tips, fast-food trends,
    Medicare or Medicaid policy, hospital operations, oncology treatment news.

Extra rules:
- Any article primarily about Medicare, Medicaid, or CMS policy → score 1.
- If you see the same event covered by multiple sources, score only the first
  occurrence normally; give all later duplicates a score of 1.
- Consumer-facing drug tips or clinical trial results → score 1 or 2.

Return ONLY a raw JSON array — no markdown fences, no extra text:
[{"index": 0, "score": 5, "reason": "one short sentence"}, ...]
"""

# Stage 2 — Formatting: only runs on articles that scored 4 or 5
FORMAT_PROMPT = """
You are an editor for the RxBenefits Intel Hub News Radar Report.
Your audience is HR directors and benefits managers at self-funded employer companies.

Reformat the article below using EXACTLY this structure:

TITLE: <A clear, specific, business-focused title in Title Case — max 15 words>
SUMMARY: <One paragraph, maximum 5 lines: what happened, why it matters to
           self-funded employers, and the key business implication>
SOURCE_LINE: Source: <Source Name>; <Publication Date>

Do not add any other text, labels, or commentary.
"""


# ==========================================
# ENVIRONMENT VALIDATION
# ==========================================
def validate_env_vars():
    print("\n🔐 Validating environment variables...", flush=True)
    required = [
        "GEMINI_API_KEY", "WP_SITE_URL", "WP_USERNAME", "WP_APP_PASS",
        "SENDER_EMAIL", "SENDER_APP_PASSWORD", "RECIPIENT_EMAIL",
    ]
    values = {
        "GEMINI_API_KEY":      GEMINI_API_KEY,
        "WP_SITE_URL":         WP_SITE_URL,
        "WP_USERNAME":         WP_USERNAME,
        "WP_APP_PASS":         WP_APP_PASS,
        "SENDER_EMAIL":        SENDER_EMAIL,
        "SENDER_APP_PASSWORD": SENDER_APP_PASSWORD,
        "RECIPIENT_EMAIL":     RECIPIENT_EMAIL,
    }
    missing = []
    for name in required:
        if values[name]:
            print(f"   ✅ {name} is set (length={len(values[name])})", flush=True)
        else:
            print(f"   ❌ {name} is NOT set or empty", flush=True)
            missing.append(name)

    if missing:
        print(f"\n❌ FATAL: Missing required env vars: {missing}", flush=True)
        return False

    print("✅ All environment variables present.\n", flush=True)
    return True


# ==========================================
# LOAD KEYWORDS & COMPANIES  ← UPDATED
# ==========================================
def _load_lines(filepath, label):
    """
    Generic loader: reads a .txt file, strips blank lines and comment
    lines (lines starting with #), and returns a list of strings.
    """
    print(f"\n📂 Looking for {label} file at: {filepath}", flush=True)

    if not os.path.exists(filepath):
        print(f"⚠️  WARNING: Cannot find '{filepath}' — skipping.", flush=True)
        return []

    lines = []
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            stripped = line.strip()
            if stripped and not stripped.startswith("#"):
                lines.append(stripped)

    print(f"✅ Loaded {len(lines)} {label}:", flush=True)
    for i, item in enumerate(lines, 1):
        print(f"   {i}. {item}", flush=True)

    return lines


def load_keywords(filepath):
    return _load_lines(filepath, "keywords")


def load_companies(filepath):           # ← NEW FUNCTION
    return _load_lines(filepath, "companies")


# ==========================================
# FETCH NEWS  (keywords + companies combined)
# ==========================================
def fetch_all_news(keywords, companies):
    """
    Fetches Google News RSS for every keyword AND every company name.
    Company searches use a more targeted query so results stay relevant.
    """
    all_articles = []

    # --- Keyword searches (broad topic terms) ---
    if keywords:
        print(f"\n📡 Fetching news for {len(keywords)} topic keywords...", flush=True)
        for i, kw in enumerate(keywords, 1):
            print(f"\n   [{i}/{len(keywords)}] Topic: '{kw}'", flush=True)
            query = f'"{kw}" when:24h'
            _fetch_rss(query, kw, "keyword", all_articles)
            time.sleep(random.uniform(2.0, 4.0))

    # --- Company searches (targeted: company name + PBM/pharmacy context) ---
    if companies:
        print(f"\n📡 Fetching news for {len(companies)} companies...", flush=True)
        for i, co in enumerate(companies, 1):
            print(f"\n   [{i}/{len(companies)}] Company: '{co}'", flush=True)
            # Adding context words keeps results relevant and avoids unrelated hits
            query = f'"{co}" pharmacy OR benefits OR PBM OR drug when:24h'
            _fetch_rss(query, co, "company", all_articles)
            time.sleep(random.uniform(2.0, 4.0))

    print(f"\n✅ Total raw articles fetched: {len(all_articles)}", flush=True)
    return all_articles


def _fetch_rss(query, source_label, search_type, all_articles):
    """Helper: fetches one RSS feed and appends results to all_articles."""
    encoded_query = urllib.parse.quote(query)
    rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=en-US&gl=US&ceid=US:en"

    try:
        feed = feedparser.parse(rss_url, request_headers=HEADERS)
        count = len(feed.entries)
        print(f"      Found {count} articles", flush=True)

        for item in feed.entries:
            source_info = item.get("source", {})
            source_name = (
                source_info.get("title", "N/A")
                if isinstance(source_info, dict)
                else getattr(source_info, "title", "N/A")
            )
            all_articles.append({
                "search_type":  search_type,   # "keyword" or "company"
                "search_term":  source_label,
                "title":        item.get("title", "N/A"),
                "link":         item.get("link", "N/A"),
                "published":    item.get("published", item.get("pubDate", "N/A")),
                "source_name":  source_name,
                "description":  item.get("summary", item.get("description", "N/A")),
            })

    except Exception as e:
        print(f"      ❌ Error fetching '{source_label}': {e}", flush=True)


# ==========================================
# DEDUPLICATION  ← NEW
# ==========================================
def deduplicate_articles(articles):
    """
    Removes duplicate articles BEFORE sending anything to AI.
    Duplicates are detected two ways:
      1. Same URL (after stripping query parameters like ?utm_source=...)
      2. Near-identical title (same first 60 characters when lowercased)
    This alone typically removes 10-20% of raw articles and saves AI tokens.
    """
    seen_urls   = set()
    seen_titles = set()
    unique      = []

    for art in articles:
        # Normalise URL — strip tracking query params
        url = art.get("link", "").split("?")[0].rstrip("/")

        # Normalise title — lowercase, remove punctuation, keep first 60 chars
        title_key = re.sub(r"[^a-z0-9 ]", "", art.get("title", "").lower())[:60].strip()

        if url in seen_urls or (title_key and title_key in seen_titles):
            continue

        seen_urls.add(url)
        if title_key:
            seen_titles.add(title_key)
        unique.append(art)

    removed = len(articles) - len(unique)
    print(f"   🧹 Deduplication: removed {removed} duplicates, {len(unique)} remain", flush=True)
    return unique


# ==========================================
# STAGE 1 — BATCH SCORING  ← NEW
# ==========================================
def score_articles_batch(articles, ai_client, batch_size=15):
    """
    Sends articles to Gemini in batches of 15 and asks it to score
    each one 1-5 for relevance.  One API call per batch instead of
    one call per article — much cheaper on tokens.

    Returns a list of tuples: (article_dict, score_int, reason_str)
    """
    scored = []
    total_batches = -(-len(articles) // batch_size)   # ceiling division

    for batch_num, batch_start in enumerate(range(0, len(articles), batch_size), 1):
        batch = articles[batch_start : batch_start + batch_size]
        print(f"\n   🔍 Scoring batch {batch_num}/{total_batches} ({len(batch)} articles)...", flush=True)

        payload = [
            {
                "index":   i,
                "title":   a["title"],
                "snippet": (a.get("description") or "")[:300],
            }
            for i, a in enumerate(batch)
        ]

        try:
            response = ai_client.models.generate_content(
                model="gemini-2.0-flash-lite",
                contents=f"{SCORING_PROMPT}\n\nArticles to score:\n{json.dumps(payload, ensure_ascii=False)}"
            )
            text = response.text.strip()

            # Strip markdown fences if Gemini adds them despite instructions
            text = re.sub(r"^```(?:json)?", "", text, flags=re.I).strip()
            text = re.sub(r"```$", "",          text).strip()

            verdicts  = json.loads(text)
            by_index  = {v["index"]: v for v in verdicts if isinstance(v, dict)}

            for i, art in enumerate(batch):
                verdict = by_index.get(i, {})
                score   = int(verdict.get("score", 3))
                reason  = str(verdict.get("reason", ""))
                scored.append((art, score, reason))
                icon = "✅" if score >= 4 else "⏭️ "
                print(f"      {icon} [{score}/5] {art['title'][:65]}", flush=True)

        except Exception as e:
            print(f"      ❌ Batch scoring failed: {e} — defaulting all to score 3", flush=True)
            for art in batch:
                scored.append((art, 3, "scoring error — kept as fallback"))

        time.sleep(random.uniform(1.0, 2.0))

    kept   = sum(1 for _, s, _ in scored if s >= 4)
    total  = len(scored)
    print(f"\n   📊 Scoring complete: {kept}/{total} articles scored ≥ 4", flush=True)
    return scored


# ==========================================
# STAGE 2 — FORMAT INDIVIDUAL ARTICLES  ← NEW
# ==========================================
def format_article(article, ai_client):
    """
    Sends a single article (that already passed scoring) to Gemini
    and asks it to rewrite the title and produce a clean summary.
    Only called for articles that scored 4 or 5 — so far fewer
    API calls than the original one-call-per-article approach.
    """
    user_prompt = f"""
Original Title:   {article['title']}
Source Name:      {article['source_name']}
Publication Date: {article['published']}
Link:             {article['link']}
Snippet:          {article.get('description', 'N/A')}
"""
    try:
        response = ai_client.models.generate_content(
            model="gemini-2.0-flash-lite",
            contents=f"{FORMAT_PROMPT}\n\n{user_prompt}"
        )
        text = response.text.strip()

        parsed = {
            "title":       "",
            "summary":     "",
            "source_line": "",
            "link":        article["link"],
            "source_name": article["source_name"],
            "date":        article["published"],
        }

        for line in text.split("\n"):
            line = line.strip()
            if line.startswith("TITLE:"):
                parsed["title"]       = line.replace("TITLE:", "").strip()
            elif line.startswith("SUMMARY:"):
                parsed["summary"]     = line.replace("SUMMARY:", "").strip()
            elif line.startswith("SOURCE_LINE:"):
                parsed["source_line"] = line.replace("SOURCE_LINE:", "").strip()

        if not parsed["title"] or not parsed["summary"] or not parsed["source_line"]:
            print(f"      ⚠️  Incomplete AI output — skipping", flush=True)
            return None

        return parsed

    except Exception as e:
        print(f"      ❌ Format error: {e}", flush=True)
        return None


# ==========================================
# SAVE & SEND  (unchanged from original)
# ==========================================
def save_excel_format(processed_articles):
    print(f"\n💾 Saving {len(processed_articles)} articles to {OUTPUT_EXCEL_FILE}", flush=True)
    try:
        os.makedirs(os.path.dirname(OUTPUT_EXCEL_FILE), exist_ok=True)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "News Radar Report"
        ws.append(["New Title", "Summary", "Source Name", "Publication Date", "Hyperlink URL"])
        for art in processed_articles:
            ws.append([art["title"], art["summary"], art["source_name"], art["date"], art["link"]])
        wb.save(OUTPUT_EXCEL_FILE)
        print("✅ Excel file saved successfully.", flush=True)
    except Exception as e:
        print(f"❌ Error saving Excel: {e}", flush=True)
        raise


def save_json_format(processed_articles):
    from datetime import datetime

    print(f"\n💾 Saving JSON to {OUTPUT_JSON_FILE}", flush=True)
    try:
        os.makedirs(os.path.dirname(OUTPUT_JSON_FILE), exist_ok=True)

        today = datetime.utcnow()
        payload = {
            "date":       today.strftime("%B %d, %Y"),
            "date_short": today.strftime("%Y-%m-%d"),
            "total":      len(processed_articles),
            "excel_url":  f"https://medtech-platform.github.io/News-Radar/data/rxbenefits_intel_hub_report_{today.strftime('%Y-%m-%d')}.xlsx",
            "articles":   processed_articles,
        }

        with open(OUTPUT_JSON_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        print(f"✅ JSON saved. Total articles: {len(processed_articles)}", flush=True)
    except Exception as e:
        print(f"❌ Error saving JSON: {e}", flush=True)
        raise


def send_email_report(processed_articles, excel_download_url):
    from email.mime.multipart import MIMEMultipart
    from datetime import datetime

    print(f"\n📧 Sending email report ({len(processed_articles)} articles)...", flush=True)
    try:
        today        = datetime.utcnow()
        subject_date = today.strftime("%m/%d/%Y")
        body_date    = today.strftime("%B %d, %Y")

        html_parts = []
        html_parts.append("""
<html><body style="font-family: Arial, sans-serif; font-size:14px;
                   color:#222; max-width:680px; margin:0 auto; padding:24px;">
""")
        html_parts.append(f"""
        <p style="margin: 16px 0;">Hi Lee Ashford,</p>
        <p style="margin: 0 0 20px 0;">
          Please find updates from <strong>{body_date}</strong> below:
        </p>
        <hr style="border:none; border-top:1px solid #ddd; margin-bottom:20px;">
        """)

        for i, art in enumerate(processed_articles, 1):
            html_parts.append(f"""
            <div style="margin-bottom:24px;">
              <div style="font-size:13px; color:#1F4E79; font-weight:500; margin-bottom:4px;">[{i}]</div>
              <div style="font-size:15px; font-weight:500; color:#000000; margin-bottom:8px; line-height:1.4;">
                {art['title']}
              </div>
              <div style="font-size:13px; color:#222; line-height:1.6; margin-bottom:8px;">
                {art['summary']}
              </div>
              <div style="font-size:12px; color:#555;">
                {art['source_line']} &nbsp;|&nbsp;
                <a href="{art['link']}" style="color:#1F4E79;">Read Full Article</a>
              </div>
            </div>
            <hr style="border:none; border-top:1px solid #eee; margin-bottom:20px;">
            """)

        html_parts.append("""
        <p style="margin-top:24px;">Regards,</p>
        <p style="font-weight:500; margin:0;">Evalueserve Team</p>
        <hr style="border:none; border-top:1px solid #ddd; margin-top:24px;">
        <div style="font-size:11px; color:#aaa; text-align:center;">
          RxBenefits Intel Hub · Daily News Radar · Automated Report
        </div>
        </body></html>
        """)

        html_body  = "".join(html_parts)

        text_parts = [f"Hi Lee Ashford,", f"Please find updates from {body_date} below:", ""]
        for i, art in enumerate(processed_articles, 1):
            text_parts.append(f"[{i}] {art['title']}")
            text_parts.append(art['summary'])
            text_parts.append(f"{art['source_line']} | Link: {art['link']}")
            text_parts.append("-" * 60)
            text_parts.append("")
        text_parts += ["Regards,", "Evalueserve Team"]
        text_body = "\n".join(text_parts)

        msg            = MIMEMultipart("alternative")
        msg["Subject"] = f"Daily News Alerts_{subject_date}"
        msg["From"]    = SENDER_EMAIL
        msg["To"]      = RECIPIENT_EMAIL

        msg.attach(MIMEText(text_body, "plain", "utf-8"))
        msg.attach(MIMEText(html_body, "html",  "utf-8"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, msg.as_string())

        print(f"✅ Email sent. Subject: Daily News Alerts_{subject_date}", flush=True)

    except Exception as e:
        print(f"❌ Email sending error: {e}", flush=True)
        raise


# ==========================================
# MAIN PIPELINE
# ==========================================
if __name__ == "__main__":
    print("=" * 60, flush=True)
    print("🚀 RxBenefits Intel Hub News Radar — Script Starting", flush=True)
    print("=" * 60, flush=True)

    try:
        # Step 1: Validate environment variables
        if not validate_env_vars():
            print("❌ Exiting due to missing environment variables.", flush=True)
            exit(1)

        # Step 2: Initialise AI client
        print("🤖 Initialising Gemini AI client...", flush=True)
        AI_CLIENT = genai.Client(api_key=GEMINI_API_KEY)
        print("✅ Gemini AI client initialised.", flush=True)

        # Step 3: Load keywords (topic terms) AND companies  ← UPDATED
        keywords  = load_keywords(INPUT_KEYWORDS_FILE)
        companies = load_companies(INPUT_COMPANIES_FILE)   # ← NEW

        if not keywords and not companies:
            print("❌ No keywords or companies loaded. Exiting.", flush=True)
            exit(1)

        # Step 4: Fetch news for both lists  ← UPDATED
        raw_articles = fetch_all_news(keywords, companies)
        if not raw_articles:
            print("⚠️  No articles fetched. Exiting.", flush=True)
            exit(0)

        # Step 4b: Deduplicate BEFORE sending to AI  ← NEW
        print(f"\n🧹 Deduplicating {len(raw_articles)} raw articles...", flush=True)
        raw_articles = deduplicate_articles(raw_articles)

        # Step 5a: Stage 1 — Score articles in batches (cheap)  ← NEW
        print(f"\n🔍 Stage 1: Scoring {len(raw_articles)} articles for relevance...", flush=True)
        scored_articles = score_articles_batch(raw_articles, AI_CLIENT, batch_size=15)

        # Keep only articles that scored 4 or 5
        candidates = [(art, score, reason) for art, score, reason in scored_articles if score >= 4]
        print(f"\n📊 {len(candidates)}/{len(raw_articles)} articles passed relevance scoring (score ≥ 4).", flush=True)

        if not candidates:
            print("⚠️  No articles passed scoring. Exiting.", flush=True)
            exit(0)

        # Step 5b: Stage 2 — Format only the articles that passed  ← NEW
        print(f"\n✍️  Stage 2: Formatting {len(candidates)} articles...", flush=True)
        processed_articles = []
        for i, (art, score, reason) in enumerate(candidates, 1):
            print(f"   [{i}/{len(candidates)}] {art['title'][:70]}", flush=True)
            res = format_article(art, AI_CLIENT)
            if res:
                processed_articles.append(res)
                print(f"      ✅ Formatted", flush=True)
            else:
                print(f"      ⚠️  Format failed — skipped", flush=True)
            time.sleep(random.uniform(0.5, 1.0))

        print(f"\n📊 Final report: {len(processed_articles)} articles.", flush=True)

        if not processed_articles:
            print("⚠️  No articles passed formatting. Exiting.", flush=True)
            exit(0)

        # Step 6: Save Excel
        save_excel_format(processed_articles)

        # Step 7: Save JSON
        save_json_format(processed_articles)

        # Step 8: Send email
        send_email_report(processed_articles, None)

        print("\n" + "=" * 60, flush=True)
        print("✅ Process completed successfully!", flush=True)
        print("=" * 60, flush=True)

    except Exception as e:
        print(f"\n❌ FATAL ERROR: {e}", flush=True)
        import traceback
        traceback.print_exc()
        exit(1)
