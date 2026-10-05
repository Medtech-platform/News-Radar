import os
import re
import time
import json
import random
import smtplib
import urllib.parse
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

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
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_KEYWORDS_FILE = os.path.join(BASE_DIR, "keywords.txt")
OUTPUT_EXCEL_FILE = os.path.join(BASE_DIR, "docs", "data", f"rxbenefits_intel_hub_report_{time.strftime('%Y-%m-%d')}.xlsx")
OUTPUT_JSON_FILE  = os.path.join(BASE_DIR, "docs", "data", f"rxbenefits_{time.strftime('%Y-%m-%d')}.json")

GEMINI_API_KEY      = os.environ.get("GEMINI_API_KEY", "").strip()
WP_SITE_URL         = os.environ.get("WP_SITE_URL", "").strip()
WP_USERNAME         = os.environ.get("WP_USERNAME", "").strip()
WP_APP_PASS         = os.environ.get("WP_APP_PASS", "").strip()
SENDER_EMAIL        = os.environ.get("SENDER_EMAIL", "").strip()
SENDER_APP_PASSWORD = os.environ.get("SENDER_APP_PASSWORD", "").strip()
RECIPIENT_EMAIL     = os.environ.get("RECIPIENT_EMAIL", "").strip()

# ==========================================
# TOKEN BUDGET — keeps Gemini calls lean
# Max articles sent to AI; cap output per article
# Gemini Flash Lite: ~1M token context, but we stay well under
# ==========================================
MAX_ARTICLES_TO_PROCESS = 60   # pre-AI cap after dedup
MAX_ARTICLES_PER_KW     = 3    # RSS results kept per keyword
GEMINI_MODEL            = "gemini-2.0-flash-lite"   # fast + cheap
GEMINI_MAX_TOKENS       = 300  # per article; title+summary+source fits in ~200

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36"
}

# ==========================================
# TIGHTENED SYSTEM PROMPT — HIGH RELEVANCE ONLY
# ==========================================
SYSTEM_PROMPT = """You are a senior intelligence analyst for RxBenefits, a pharmacy benefits management (PBM) company that exclusively serves self-funded employer health plans.

Your task: evaluate ONE news article and decide whether it is HIGH-RELEVANCE to RxBenefits.

─── WHAT RxBenefits CARES ABOUT ───────────────────────────────
• PBM industry: contracts, pricing models, audits, reform, litigation, M&A
• Self-funded employer health plans: cost management, benefits strategy, plan design
• Drug pricing: rebates, spread pricing, pass-through, net vs gross spend, inflation
• Specialty pharmacy & drugs: coverage, costs, prior auth, step therapy
• GLP-1 / obesity drugs: employer coverage decisions, cost impact
• Biosimilars: uptake, savings, formulary changes
• 340B Program: policy changes, litigation, reform
• PBM transparency: FTC actions, Congressional actions, state legislation
• Pharmacy coalitions and GPOs for employers
• Health plan and commercial insurance cost trends

─── HARD EXCLUDE (respond SKIP immediately) ────────────────────
• Medicare, Medicaid, CMS, ACA marketplace, Medigap, CHIP
• International/non-US news (UK, EU, Canada, India, etc.)
• Clinical trial results, new drug approvals (unless pricing/coverage angle)
• Consumer personal finance or individual insurance shopping
• Hospital/physician fee-for-service topics not connected to PBM
• Generic news about "healthcare" with no pharmacy benefit angle
• Duplicate or near-duplicate of another article already in today's batch

─── RELEVANCE TEST ─────────────────────────────────────────────
Ask: "Would a self-funded employer or PBM executive find this directly actionable or strategically important TODAY?"
If NO → SKIP
If YES → process it

─── OUTPUT FORMAT ──────────────────────────────────────────────
If article does NOT pass → respond ONLY with: SKIP

If article PASSES → respond EXACTLY in this format (no extra text, no markdown):
TITLE: <Punchy, specific title in Title Case — max 15 words>
SUMMARY: <2-3 sentences: what happened, why it matters to self-funded employers / PBMs, key implication>
SOURCE_LINE: Source: <Source Name>; <Publication Date>"""


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


def load_keywords(filepath):
    print(f"\n📂 Loading keywords from: {filepath}", flush=True)
    if not os.path.exists(filepath):
        print(f"❌ ERROR: Cannot find '{filepath}'!", flush=True)
        return []
    with open(filepath, "r", encoding="utf-8") as f:
        keywords = [line.strip() for line in f if line.strip() and not line.startswith("#")]
    print(f"✅ Loaded {len(keywords)} keywords.", flush=True)
    return keywords


def load_companies(filepath):
    filepath = os.path.join(BASE_DIR, "Companies.txt")
    if not os.path.exists(filepath):
        print(f"⚠️  Companies.txt not found — skipping competitor news", flush=True)
        return []
    with open(filepath, "r", encoding="utf-8") as f:
        companies = [line.strip() for line in f if line.strip() and not line.startswith("#")]
    print(f"✅ Loaded {len(companies)} companies.", flush=True)
    return companies


def fetch_all_news(keywords, companies):
    all_articles = []
    print(f"\n📡 Fetching news for {len(keywords)} keywords...", flush=True)

    for i, keyword in enumerate(keywords, 1):
        print(f"\n   [{i}/{len(keywords)}] Fetching: '{keyword}'", flush=True)
        query = f'"{keyword}" when:24h'
        encoded_query = urllib.parse.quote(query)
        rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=en-US&gl=US&ceid=US:en"
        try:
            feed = feedparser.parse(rss_url, request_headers=HEADERS)
            count = len(feed.entries)
            print(f"      Found {count} articles", flush=True)
            for item in feed.entries[:MAX_ARTICLES_PER_KW]:
                source_info = item.get("source", {})
                source_name = (
                    source_info.get("title", "N/A")
                    if isinstance(source_info, dict)
                    else getattr(source_info, "title", "N/A")
                )
                all_articles.append({
                    "keyword":     keyword,
                    "title":       item.get("title", "N/A"),
                    "link":        item.get("link", "N/A"),
                    "published":   item.get("published", item.get("pubDate", "N/A")),
                    "source_name": source_name,
                    "description": item.get("summary", item.get("description", "N/A")),
                })
        except Exception as e:
            print(f"      ❌ Error fetching '{keyword}': {e}", flush=True)
        time.sleep(random.uniform(1.0, 2.0))

    # Competitor / company news
    if companies:
        print(f"\n📡 Fetching competitor news for {len(companies)} companies...", flush=True)
        for i, company in enumerate(companies, 1):
            print(f"\n   [{i}/{len(companies)}] Fetching: '{company}'", flush=True)
            query = f'"{company}" pharmacy benefits when:24h'
            encoded_query = urllib.parse.quote(query)
            rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=en-US&gl=US&ceid=US:en"
            try:
                feed = feedparser.parse(rss_url, request_headers=HEADERS)
                count = len(feed.entries)
                print(f"      Found {count} articles", flush=True)
                for item in feed.entries[:2]:   # cap at 2 per company
                    source_info = item.get("source", {})
                    source_name = (
                        source_info.get("title", "N/A")
                        if isinstance(source_info, dict)
                        else getattr(source_info, "title", "N/A")
                    )
                    all_articles.append({
                        "keyword":     company,
                        "title":       item.get("title", "N/A"),
                        "link":        item.get("link", "N/A"),
                        "published":   item.get("published", item.get("pubDate", "N/A")),
                        "source_name": source_name,
                        "description": item.get("summary", item.get("description", "N/A")),
                    })
            except Exception as e:
                print(f"      ❌ Error fetching '{company}': {e}", flush=True)
            time.sleep(random.uniform(1.5, 3.0))

    return all_articles


def deduplicate_articles(articles):
    """Remove exact URL duplicates and near-duplicate titles before AI processing."""
    seen_urls   = set()
    seen_titles = set()
    unique      = []
    for art in articles:
        url = art.get("link", "").split("?")[0].rstrip("/")
        title_key = re.sub(r"[^a-z0-9 ]", "", art.get("title", "").lower())[:80].strip()
        if url in seen_urls or title_key in seen_titles:
            continue
        seen_urls.add(url)
        if title_key:
            seen_titles.add(title_key)
        unique.append(art)
    removed = len(articles) - len(unique)
    print(f"   🧹 Deduplication: removed {removed} duplicates, {len(unique)} remain", flush=True)
    return unique


def process_article_with_ai(article, ai_client):
    # Keep snippet short to conserve tokens — first 400 chars is enough for relevance
    snippet = (article.get("description") or "")[:400]

    user_prompt = f"""Evaluate this article for RxBenefits relevance:

Original Title: {article['title'][:200]}
Source Name: {article['source_name']}
Publication Date: {article['published']}
Link: {article['link']}
Snippet: {snippet}"""

    try:
        response = ai_client.models.generate_content(
            model=GEMINI_MODEL,
            contents=f"{SYSTEM_PROMPT}\n\n{user_prompt}",
            config={"max_output_tokens": GEMINI_MAX_TOKENS}
        )
        text = response.text.strip() if response.text else ""

        if not text or text.upper().startswith("SKIP"):
            return None

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
                parsed["title"] = line.replace("TITLE:", "", 1).strip()
            elif line.startswith("SUMMARY:"):
                parsed["summary"] = line.replace("SUMMARY:", "", 1).strip()
            elif line.startswith("SOURCE_LINE:"):
                parsed["source_line"] = line.replace("SOURCE_LINE:", "", 1).strip()

        if not parsed["title"] or not parsed["summary"] or not parsed["source_line"]:
            print(f"      ⚠️  Incomplete AI output — skipping", flush=True)
            return None

        return parsed

    except Exception as e:
        print(f"      ❌ AI Processing Error: {e}", flush=True)
        return None


def save_excel_format(processed_articles):
    print(f"\n💾 Saving {len(processed_articles)} articles to Excel...", flush=True)
    try:
        os.makedirs(os.path.dirname(OUTPUT_EXCEL_FILE), exist_ok=True)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "News Radar Report"

        # Header row styling
        headers = ["#", "Title", "Summary", "Source", "Publication Date", "Link"]
        header_fill = PatternFill("solid", fgColor="1F4E79")
        header_font = Font(bold=True, color="FFFFFF")
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(wrap_text=True, vertical="center")

        for i, art in enumerate(processed_articles, 1):
            ws.append([i, art["title"], art["summary"], art["source_name"], art["date"], art["link"]])

        # Column widths
        col_widths = [4, 35, 60, 20, 20, 40]
        for col, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w

        # Wrap text for all data rows
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        wb.save(OUTPUT_EXCEL_FILE)
        print("✅ Excel file saved successfully.", flush=True)
    except Exception as e:
        print(f"❌ Error saving Excel: {e}", flush=True)
        raise


def save_json_format(processed_articles):
    print(f"\n💾 Saving JSON...", flush=True)
    try:
        os.makedirs(os.path.dirname(OUTPUT_JSON_FILE), exist_ok=True)
        today = datetime.utcnow()
        payload = {
            "date":       today.strftime("%B %d, %Y"),
            "date_short": today.strftime("%Y-%m-%d"),
            "total":      len(processed_articles),
            "excel_url":  f"https://medtech-platform.github.io/News-Radar/data/rxbenefits_intel_hub_report_{today.strftime('%Y-%m-%d')}.xlsx",
            "articles":   processed_articles
        }
        with open(OUTPUT_JSON_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print(f"✅ JSON saved. Total articles: {len(processed_articles)}", flush=True)
    except Exception as e:
        print(f"❌ Error saving JSON: {e}", flush=True)
        raise


def send_email_report(processed_articles):
    """
    Send email via Gmail using SMTP with STARTTLS on port 587.

    FIX NOTES:
    - Uses port 587 + STARTTLS (NOT port 465 + SSL).
      GitHub Actions blocks outbound port 465, causing SMTPServerDisconnected.
    - Increased timeout to 60s to handle slow GitHub runner network.
    - Added explicit server.ehlo() calls before and after starttls().
    - SENDER_APP_PASSWORD must be a Gmail App Password (16 chars, no spaces),
      NOT your regular Gmail password. 2FA must be enabled on the sender account.

    Gmail setup:
      1. Enable 2-Factor Authentication on the sender Gmail account
      2. Go to: Google Account → Security → 2-Step Verification → App Passwords
      3. Create an App Password for "Mail" / "Other"
      4. Copy the 16-character code (no spaces) → store as SENDER_APP_PASSWORD secret
    """
    print(f"\n📧 Sending email report ({len(processed_articles)} articles)...", flush=True)
    try:
        today        = datetime.utcnow()
        subject_date = today.strftime("%m/%d/%Y")
        body_date    = today.strftime("%B %d, %Y")

        # ---- HTML body ----
        html_parts = ["""
<html><body style="font-family: Arial, sans-serif; font-size:14px;
                   color:#222; max-width:680px; margin:0 auto; padding:24px;">
"""]
        html_parts.append(f"""
        <p style="margin: 16px 0;">Hi Lee Ashford,</p>
        <p style="margin: 0 0 20px 0;">
          Please find <strong>{len(processed_articles)} high-relevance updates</strong>
          from <strong>{body_date}</strong> below:
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
        html_body = "".join(html_parts)

        # ---- Plain text fallback ----
        text_parts = [
            f"Hi Lee Ashford,",
            f"Please find {len(processed_articles)} high-relevance updates from {body_date} below:",
            "",
        ]
        for i, art in enumerate(processed_articles, 1):
            text_parts += [
                f"[{i}] {art['title']}",
                art['summary'],
                f"{art['source_line']} | Link: {art['link']}",
                "-" * 60,
                "",
            ]
        text_parts += ["Regards,", "Evalueserve Team"]
        text_body = "\n".join(text_parts)

        # ---- Build message ----
        msg            = MIMEMultipart("alternative")
        msg["Subject"] = f"Daily News Alerts_{subject_date}"
        msg["From"]    = SENDER_EMAIL
        msg["To"]      = RECIPIENT_EMAIL
        msg.attach(MIMEText(text_body, "plain", "utf-8"))
        msg.attach(MIMEText(html_body, "html",  "utf-8"))

        # ---- Send via STARTTLS on port 587 ----
        # FIX: timeout=60 (was 30) to survive slow GitHub Actions network
        print("   📬 Connecting to smtp.gmail.com:587 (STARTTLS)...", flush=True)
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=60) as server:
            server.ehlo()           # identify ourselves to the server
            server.starttls()       # upgrade to encrypted connection
            server.ehlo()           # re-identify over TLS (required by Gmail)
            print("   🔐 Logging in...", flush=True)
            server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
            print("   📤 Sending message...", flush=True)
            server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, msg.as_string())

        print(f"✅ Email sent successfully. Subject: Daily News Alerts_{subject_date}", flush=True)

    except smtplib.SMTPAuthenticationError as e:
        print(f"❌ Email AUTH FAILED: {e}", flush=True)
        print("   → Check: Is SENDER_APP_PASSWORD a Gmail App Password (not your real password)?", flush=True)
        print("   → Check: Is 2-Factor Auth enabled on the sender Gmail account?", flush=True)
        raise
    except smtplib.SMTPServerDisconnected as e:
        print(f"❌ SMTP connection dropped: {e}", flush=True)
        print("   → This usually means port 587 is blocked or App Password is wrong.", flush=True)
        raise
    except Exception as e:
        print(f"❌ Email sending error: {e}", flush=True)
        raise


# ==========================================
# MAIN
# ==========================================
if __name__ == "__main__":
    print("=" * 60, flush=True)
    print("🚀 RxBenefits Intel Hub News Radar — Script Starting", flush=True)
    print("=" * 60, flush=True)

    try:
        # Step 1: Validate env vars
        if not validate_env_vars():
            print("❌ Exiting due to missing environment variables.", flush=True)
            exit(1)

        # Step 2: Initialize Gemini AI client
        print("🤖 Initializing Gemini AI client...", flush=True)
        AI_CLIENT = genai.Client(api_key=GEMINI_API_KEY)
        print("✅ Gemini AI client initialized.", flush=True)

        # Step 3: Load keywords + companies
        keywords  = load_keywords(INPUT_KEYWORDS_FILE)
        companies = load_companies(os.path.join(BASE_DIR, "Companies.txt"))
        if not keywords:
            print("❌ No keywords loaded. Exiting.", flush=True)
            exit(1)

        # Step 4: Fetch news
        raw_articles = fetch_all_news(keywords, companies)
        if not raw_articles:
            print("⚠️  No articles fetched. Exiting.", flush=True)
            exit(0)

        print(f"\n📰 Total raw articles fetched: {len(raw_articles)}", flush=True)

        # Step 4b: Deduplicate before AI (saves tokens)
        raw_articles = deduplicate_articles(raw_articles)

        # Step 4c: Cap articles sent to Gemini to stay within token budget
        if len(raw_articles) > MAX_ARTICLES_TO_PROCESS:
            print(f"   ✂️  Capping at {MAX_ARTICLES_TO_PROCESS} articles to manage Gemini token budget", flush=True)
            raw_articles = raw_articles[:MAX_ARTICLES_TO_PROCESS]

        # Step 5: Process with AI — high-relevance filter
        print(f"\n🤖 Processing {len(raw_articles)} articles with Gemini AI (high-relevance filter)...", flush=True)
        processed_articles = []
        for i, art in enumerate(raw_articles, 1):
            title_preview = art["title"][:70]
            print(f"   [{i}/{len(raw_articles)}] {title_preview}", flush=True)
            res = process_article_with_ai(art, AI_CLIENT)
            if res:
                processed_articles.append(res)
                print(f"      ✅ Kept", flush=True)
            else:
                print(f"      ⏭️  Skipped", flush=True)
            time.sleep(random.uniform(0.3, 0.8))   # reduced delay — Flash Lite handles rate limits well

        print(f"\n📊 {len(processed_articles)}/{len(raw_articles)} articles kept after AI filter.", flush=True)

        if not processed_articles:
            print("⚠️  No articles passed AI filter. Exiting.", flush=True)
            exit(0)

        # Step 6: Save Excel
        save_excel_format(processed_articles)

        # Step 7: Save JSON
        save_json_format(processed_articles)

        # Step 8: Send email
        send_email_report(processed_articles)

        print("\n" + "=" * 60, flush=True)
        print("✅ Process completed successfully!", flush=True)
        print("=" * 60, flush=True)

    except Exception as e:
        print(f"\n❌ FATAL ERROR: {e}", flush=True)
        import traceback
        traceback.print_exc()
        exit(1)
