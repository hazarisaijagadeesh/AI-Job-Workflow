"""
H1B Job Hunt Automation — Sai Jagadeesh Hazari
Runs every 12 hours via GitHub Actions (or cron).
Scrapes company career pages → filters H1B → ATS scores vs resume → saves Excel → emails top matches.
"""

import os, re, json, time, logging, smtplib, hashlib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# CONFIGURATION — edit these or use env vars
# ─────────────────────────────────────────────

YOUR_NAME      = os.getenv("YOUR_NAME",  "Sai Jagadeesh Hazari")
YOUR_EMAIL     = os.getenv("YOUR_EMAIL", "hazarisaijagadeesh@gmail.com")
GMAIL_APP_PASS = os.getenv("GMAIL_APP_PASS", "")       # Gmail App Password (not your real password)
EXCEL_PATH     = Path("output/job_tracker.xlsx")
SEEN_JOBS_FILE = Path("output/seen_jobs.json")
MIN_ATS_SCORE  = 70   # only shortlist jobs scoring >= this
MAX_APPLY_PER_RUN = 5 # cap applications per run to avoid spam

RESUME_SUMMARY = """
Senior SDET / QA Automation Engineer with 4+ years at American Express.
Skills: Selenium WebDriver, Playwright, Java, Python, JavaScript, RestAssured,
Cucumber BDD, TestNG, JUnit, Postman, PyTest, Jenkins, GitHub Actions, GitLab CI,
Maven, Docker, AWS (EC2, S3, RDS), MySQL, MongoDB, Jira, Agile/Scrum, Salesforce.
Led migration of 1000+ test scenarios from Selenium to Playwright (~60% faster).
Increased regression coverage by 70%. Automated 500+ end-to-end web/API scenarios.
Delivered zero critical defects across multiple production releases.
Built AI-powered Playwright test script generator from Rally user stories.
MS Computer Science, University of Central Missouri.
Authorized to work in the US (requires H1B visa sponsorship).
"""

# Target companies: (display name, careers page URL, optional keyword to find job links)
# ── Direct company career pages ──────────────────────────────────────────────
COMPANIES = [
    ("Google",          "https://careers.google.com/jobs/results/?q=QA+automation&employment_type=FULL_TIME",  None),
    ("Meta",            "https://www.metacareers.com/jobs?q=QA+automation&teams[0]=Engineering",                None),
    ("Amazon",          "https://www.amazon.jobs/en/search?base_query=SDET+QA+automation&loc_query=",          None),
    ("Microsoft",       "https://jobs.microsoft.com/us/en/search#q=QA%20automation%20SDET&p=1",               None),
    ("Salesforce",      "https://careers.salesforce.com/en/jobs/?search=QA+automation&department=Software+Engineering", None),
    ("JPMorgan Chase",  "https://jpmc.fa.oraclecloud.com/hcmUI/CandidateExperience/en/sites/CX_1001/requisitions?keyword=SDET+QA+automation", None),
    ("Apple",           "https://jobs.apple.com/en-us/search?search=QA+automation+SDET&sort=relevance",        None),
    ("Netflix",         "https://jobs.netflix.com/search?q=QA%20automation",                                  None),
    ("Stripe",          "https://stripe.com/jobs/search?query=QA+automation",                                 None),
    ("Uber",            "https://www.uber.com/us/en/careers/list/?query=QA+automation",                       None),
    ("Twilio",          "https://boards.greenhouse.io/twilio",                                                 "QA"),
    ("Atlassian",       "https://www.atlassian.com/company/careers/all-jobs?team=Engineering&search=QA",       None),
    ("ServiceNow",      "https://careers.servicenow.com/careers/jobs?keywords=QA+automation+SDET",             None),
    ("Workday",         "https://www.workday.com/en-us/company/careers/open-positions.html?q=QA+automation",   None),
    ("Adobe",           "https://careers.adobe.com/us/en/search-results?keywords=QA+automation",               None),
    ("Intuit",          "https://jobs.intuit.com/search-jobs?keyword=SDET+QA+automation",                      None),
    ("PayPal",          "https://careers.pypl.com/jobs/?keyword=QA+automation+SDET",                           None),
    ("Cisco",           "https://jobs.cisco.com/jobs/SearchJobs/QA%20automation%20SDET",                       None),
    ("Oracle",          "https://careers.oracle.com/jobs/#en/sites/jobsearch/jobs?keyword=QA+automation+SDET", None),
    ("IBM",             "https://www.ibm.com/employment/#jobs?job-search=QA+automation+SDET",                  None),
]

# ── Job portals — parsed with dedicated scrapers below ───────────────────────
# Each entry: (portal_name, search_url, parser_key)
JOB_PORTALS = [
    # LinkedIn public job search (no login needed for listings page)
    ("LinkedIn",   "https://www.linkedin.com/jobs/search/?keywords=SDET+QA+automation&location=United+States&f_WT=2&f_JT=F",  "linkedin"),
    # Indeed
    ("Indeed",     "https://www.indeed.com/jobs?q=SDET+QA+automation+%22visa+sponsorship%22&l=United+States&jt=fulltime",     "indeed"),
    # Dice — tech-focused, great for SDET roles
    ("Dice",       "https://www.dice.com/jobs?q=SDET+QA+automation&location=United+States&filters.workplaceTypes=Remote&filters.employmentType=FULLTIME", "dice"),
    # Monster
    ("Monster",    "https://www.monster.com/jobs/search?q=SDET+QA+automation&where=United+States&jobtype=fulltime",            "monster"),
    # ZipRecruiter
    ("ZipRecruiter","https://www.ziprecruiter.com/Jobs/SDET-QA-Automation?radius=25&days=3",                                   "generic"),
    # SimplyHired
    ("SimplyHired","https://www.simplyhired.com/search?q=SDET+QA+automation+visa+sponsorship&l=United+States",                 "simplyhired"),
    # Glassdoor
    ("Glassdoor",  "https://www.glassdoor.com/Job/united-states-sdet-qa-automation-jobs-SRCH_IL.0,13_IN1_KO14,32.htm",        "glassdoor"),
    # CareerBuilder
    ("CareerBuilder","https://www.careerbuilder.com/jobs?keywords=SDET+QA+automation&location=United+States&emp=jtft",        "generic"),
    # Wellfound (AngelList) — great for startup SDET roles with sponsorship
    ("Wellfound",  "https://wellfound.com/jobs?role=qa-engineer&remote=true",                                                  "generic"),
    # Greenhouse job board aggregator
    ("Greenhouse", "https://boards.greenhouse.io/embed/job_board?for=",                                                        "greenhouse"),
    # Lever job board aggregator
    ("Lever",      "https://jobs.lever.co/",                                                                                   "lever"),
    # Built In — tech jobs with H1B filter
    ("Built In",   "https://builtin.com/jobs/dev-engineer/qa?title=QA+Automation+SDET",                                       "generic"),
    # Remotive — remote tech jobs
    ("Remotive",   "https://remotive.com/remote-jobs/qa?search=automation",                                                    "remotive"),
    # We Work Remotely
    ("WeWorkRemotely", "https://weworkremotely.com/remote-jobs/search?term=QA+automation+SDET",                               "generic"),
]

H1B_POSITIVE = [
    "visa sponsorship", "sponsor visa", "h1b", "h-1b", "will sponsor",
    "work authorization provided", "we sponsor", "sponsorship available",
    "immigration sponsorship", "visa support", "relocation and visa"
]
H1B_NEGATIVE = [
    "no sponsorship", "not sponsor", "must be authorized", "must be legally authorized",
    "no visa", "citizens only", "us citizens and permanent residents only",
    "must have authorization to work", "no h1b"
]
USA_PATTERNS = [
    r"\b(?:united states|united states of america|usa|u\.s\.a|u\.s\.|us|america)\b",
    r"\b(?:new york|california|texas|florida|illinois|washington|seattle|san francisco|ny|ca|tx)\b",
]

INDIA_PATTERNS = [
    r"\b(?:india|indian)\b",
    r"\b(?:chennai|bangalore|bangaluru|mumbai|delhi|hyderabad|pune|kolkata|gurgaon|noida)\b",
]

# allow remote only when explicitly mentioning US/India
REMOTE_OK = [r"\bremote\b.*\b(?:usa|us|united states|india|indian)\b", r"\b(?:usa|india)\b.*\bremote\b"]

ROLE_POSITIVE = [
    "sdet", "software engineer in test", "qa automation", "test automation",
    "quality engineer", "automation engineer", "qa engineer", "quality assurance",
    "test engineer"
]

ROLE_NEGATIVE = [
    "data scientist", "sales", "recruiter", "marketing", "human resources",
    "hr", "accountant", "product manager", "business analyst"
]

def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip().lower())


def location_allowed(job: dict) -> bool:
    text = " ".join([
        job.get("title", ""), job.get("company", ""), job.get("snippet", ""),
        job.get("source", ""), job.get("url", ""), job.get("location", "")
    ])
    text = normalize_text(text)

    # explicit country match
    if any(re.search(p, text, re.I) for p in USA_PATTERNS):
        return True
    if any(re.search(p, text, re.I) for p in INDIA_PATTERNS):
        return True

    # remote is allowed only if it explicitly mentions US/India
    if any(re.search(p, text, re.I) for p in REMOTE_OK):
        return True

    return False


def role_allowed(job: dict) -> bool:
    """Return True if job title/snippet matches target roles (QA/SDET/test automation).
    Also excludes clearly unrelated roles via ROLE_NEGATIVE."""
    text = " ".join([job.get("title", ""), job.get("snippet", ""), job.get("company", "")])
    text = normalize_text(text)
    if any(neg in text for neg in ROLE_NEGATIVE):
        return False
    return any(p in text for p in ROLE_POSITIVE)
# ─────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────

def job_id(title: str, company: str) -> str:
    return hashlib.md5(f"{title.lower().strip()}{company.lower().strip()}".encode()).hexdigest()[:12]

def load_seen() -> set:
    if SEEN_JOBS_FILE.exists():
        return set(json.loads(SEEN_JOBS_FILE.read_text()))
    return set()

def save_seen(seen: set):
    SEEN_JOBS_FILE.write_text(json.dumps(list(seen)))

def h1b_status(text: str) -> str:
    """Returns 'yes', 'no', or 'unknown' based on job description text."""
    t = text.lower()
    if any(p in t for p in H1B_NEGATIVE):
        return "no"
    if any(p in t for p in H1B_POSITIVE):
        return "yes"
    return "unknown"


def annotate_job(job: dict) -> dict:
    """Add derived fields such as H1B sponsorship likelihood."""
    text = " ".join([
        job.get("title", ""), job.get("snippet", ""), job.get("company", ""),
        job.get("source", ""), job.get("url", "")
    ])
    job["h1b_likely"] = h1b_status(text)
    return job

# ─────────────────────────────────────────────
# SCRAPING
# ─────────────────────────────────────────────

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
}

def fetch_page(url: str, timeout=15) -> str | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout)
        r.raise_for_status()
        return r.text
    except Exception as e:
        log.warning(f"Fetch failed {url}: {e}")
        return None

def extract_jobs_from_html(html: str, company: str) -> list[dict]:
    """
    Generic extractor — looks for common job listing patterns.
    Returns list of {title, url, snippet} dicts.
    """
    soup = BeautifulSoup(html, "html.parser")
    jobs = []

    # Remove nav/footer noise
    for tag in soup(["nav", "footer", "header", "script", "style"]):
        tag.decompose()

    # Strategy 1: look for <a> tags containing job-like text near role keywords
    role_keywords = r"(engineer|developer|sdet|qa|quality|automation|test|analyst)"
    seen_titles = set()

    for a in soup.find_all("a", href=True):
        title = a.get_text(strip=True)
        if not title or len(title) < 6 or len(title) > 120:
            continue
        if not re.search(role_keywords, title, re.I):
            continue
        if title in seen_titles:
            continue
        seen_titles.add(title)
        href = a["href"]
        if href.startswith("/"):
            href = f"https://{requests.utils.urlparse(a.base_url if hasattr(a,'base_url') else '').netloc}{href}" if False else href
        jobs.append({"title": title, "url": href, "company": company, "snippet": ""})

    # Strategy 2: JSON-LD structured data (many modern career sites use this)
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
            if isinstance(data, list):
                items = data
            elif isinstance(data, dict):
                items = data.get("itemListElement", [data])
            else:
                continue
            for item in items:
                if item.get("@type") in ("JobPosting", "ListItem"):
                    title = item.get("title") or item.get("name", "")
                    url   = item.get("url", "")
                    desc  = item.get("description", "")[:500]
                    if title and re.search(role_keywords, title, re.I):
                        jobs.append({"title": title, "url": url, "company": company, "snippet": desc})
        except Exception:
            pass

    return jobs[:30]  # cap per company


# ─────────────────────────────────────────────
# PORTAL-SPECIFIC SCRAPERS
# ─────────────────────────────────────────────

def scrape_linkedin(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("div.base-card, li.jobs-search__results-list > div, div.job-search-card"):
        title_el = card.select_one("h3.base-search-card__title, h3, .job-title")
        co_el    = card.select_one("h4.base-search-card__subtitle, h4, .job-listing-name")
        link_el  = card.select_one("a.base-card__full-link, a[href*='/jobs/view/']")
        if not title_el: continue
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "LinkedIn listing",
            "url":     link_el["href"] if link_el else "",
            "snippet": "",
            "source":  "LinkedIn",
        })
    return jobs[:40]

def scrape_indeed(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    # Indeed uses data-jk attributes and mosaic components
    for card in soup.select("div.job_seen_beacon, div.jobsearch-SerpJobCard, li[data-jk]"):
        title_el = card.select_one("h2.jobTitle span, a.jobtitle, [data-testid='job-title']")
        co_el    = card.select_one("span.companyName, .company, [data-testid='company-name']")
        link_el  = card.select_one("a[id^='job_'], a[href*='/viewjob'], a[href*='/rc/clk']")
        snippet_el = card.select_one("div.job-snippet, .summary")
        if not title_el: continue
        href = ""
        if link_el:
            href = link_el.get("href","")
            if href.startswith("/"):
                href = "https://www.indeed.com" + href
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "Indeed listing",
            "url":     href,
            "snippet": snippet_el.get_text(strip=True)[:400] if snippet_el else "",
            "source":  "Indeed",
        })
    # Also try JSON embedded data
    for script in soup.find_all("script", type="application/json"):
        try:
            data = json.loads(script.string or "")
            hits = data.get("props",{}).get("pageProps",{}).get("jobResults",{}).get("results",[])
            for h in hits[:20]:
                t = h.get("displayTitle") or h.get("title","")
                if t:
                    jobs.append({"title":t,"company":h.get("company",""),"url":h.get("viewJobLink",""),"snippet":h.get("snippet","")[:400],"source":"Indeed"})
        except Exception:
            pass
    return jobs[:40]

def scrape_dice(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    # Dice embeds job data as JSON in a __NEXT_DATA__ script
    script = soup.find("script", id="__NEXT_DATA__")
    if script:
        try:
            data = json.loads(script.string)
            results = (data.get("props",{}).get("pageProps",{})
                          .get("initialState",{}).get("search",{})
                          .get("results",[]))
            for r in results:
                jobs.append({
                    "title":   r.get("title",""),
                    "company": r.get("employerDisplayName", r.get("companyPageUrl","Dice listing")),
                    "url":     "https://www.dice.com/jobs/" + r.get("id","") if r.get("id") else r.get("applyUrls",[""])[0],
                    "snippet": r.get("jobDescription","")[:400],
                    "source":  "Dice",
                })
        except Exception as e:
            log.debug(f"Dice JSON parse: {e}")
    # Fallback: HTML cards
    if not jobs:
        for card in soup.select("dhi-search-result, div[data-cy='search-result']"):
            title_el = card.select_one("a.card-title-link, h5")
            co_el    = card.select_one("span.employer-name, .company-name")
            if title_el:
                jobs.append({"title": title_el.get_text(strip=True),
                             "company": co_el.get_text(strip=True) if co_el else "Dice listing",
                             "url": title_el.get("href",""), "snippet":"","source":"Dice"})
    return jobs[:40]

def scrape_monster(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("section.card-content, div.summary, article[data-jobid]"):
        title_el = card.select_one("h2.title a, a.job-title, h3.title")
        co_el    = card.select_one("div.company, span.name, div.company-name")
        snippet_el = card.select_one("div.job-description, p.summary-text")
        if not title_el: continue
        href = title_el.get("href","")
        if href.startswith("/"):
            href = "https://www.monster.com" + href
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "Monster listing",
            "url":     href,
            "snippet": snippet_el.get_text(strip=True)[:400] if snippet_el else "",
            "source":  "Monster",
        })
    return jobs[:40]

def scrape_simplyhired(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("article.SerpJob, div[data-testid='JobCard']"):
        title_el = card.select_one("h2 a, a.jobposting-title, h3 a")
        co_el    = card.select_one("span[data-testid='company'], .company, b")
        snippet_el = card.select_one("p.jobposting-snippet, div.SerpJob-jobExcerpt")
        if not title_el: continue
        href = title_el.get("href","")
        if href.startswith("/"):
            href = "https://www.simplyhired.com" + href
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "SimplyHired listing",
            "url":     href,
            "snippet": snippet_el.get_text(strip=True)[:400] if snippet_el else "",
            "source":  "SimplyHired",
        })
    return jobs[:40]

def scrape_glassdoor(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("li.react-job-listing, div.jobCard, article"):
        title_el = card.select_one("a[data-test='job-title'], span.job-title, h2")
        co_el    = card.select_one("div.employer-name, span.employer-short-name, .companyName")
        if not title_el: continue
        href = title_el.get("href","") if title_el.name == "a" else ""
        if href.startswith("/"):
            href = "https://www.glassdoor.com" + href
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "Glassdoor listing",
            "url":     href,
            "snippet": "",
            "source":  "Glassdoor",
        })
    return jobs[:40]

def scrape_remotive(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for card in soup.select("li.job, div.job-card, section[itemtype*='JobPosting']"):
        title_el = card.select_one("h2, h3, span[itemprop='title'], a.job-title")
        co_el    = card.select_one("span[itemprop='name'], .company-name, strong")
        link_el  = card.select_one("a[href*='/remote-jobs/']")
        if not title_el: continue
        href = link_el["href"] if link_el else ""
        if href and not href.startswith("http"):
            href = "https://remotive.com" + href
        jobs.append({
            "title":   title_el.get_text(strip=True),
            "company": co_el.get_text(strip=True) if co_el else "Remotive listing",
            "url":     href,
            "snippet": "",
            "source":  "Remotive",
        })
    return jobs[:40]

PORTAL_PARSERS = {
    "linkedin":    scrape_linkedin,
    "indeed":      scrape_indeed,
    "dice":        scrape_dice,
    "monster":     scrape_monster,
    "simplyhired": scrape_simplyhired,
    "glassdoor":   scrape_glassdoor,
    "remotive":    scrape_remotive,
    "generic":     lambda html: [],   # falls back to generic extractor
    "greenhouse":  lambda html: [],
    "lever":       lambda html: [],
}

def scrape_portals() -> list[dict]:
    all_jobs = []
    for portal_name, url, parser_key in JOB_PORTALS:
        log.info(f"Scraping portal: {portal_name}...")
        html = fetch_page(url)
        if not html:
            log.warning(f"  Could not fetch {portal_name}")
            continue
        parser = PORTAL_PARSERS.get(parser_key, PORTAL_PARSERS["generic"])
        jobs = parser(html)
        # Fallback to generic extractor if portal parser found nothing
        if not jobs:
            jobs = extract_jobs_from_html(html, portal_name)
        # Tag source portal
        for j in jobs:
            j.setdefault("source", portal_name)
            j.setdefault("company", portal_name + " listing")
        log.info(f"  → {len(jobs)} listings from {portal_name}")
        all_jobs.extend(jobs)
        time.sleep(3)
    return all_jobs


def scrape_all_companies() -> list[dict]:
    all_jobs = []
    # 1. Direct company pages
    log.info("── Scraping direct company career pages ──")
    for company, url, keyword in COMPANIES:
        log.info(f"Scraping {company}...")
        html = fetch_page(url)
        if not html:
            continue
        jobs = extract_jobs_from_html(html, company)
        for j in jobs:
            j.setdefault("source", "Company Site")
        log.info(f"  → found {len(jobs)} potential listings")
        all_jobs.extend(jobs)
        time.sleep(2)
    # 2. Job portals
    log.info("── Scraping job portals ──")
    all_jobs.extend(scrape_portals())
    log.info(f"Total raw listings: {len(all_jobs)}")
    return all_jobs


# ─────────────────────────────────────────────
# ATS SCORING — 100% FREE, local keyword match
# ─────────────────────────────────────────────

SKILL_WEIGHTS = [
    ("playwright",        12), ("selenium",          12), ("sdet",              10),
    ("qa automation",     10), ("test automation",    9),  ("java",               7),
    ("python",             6), ("javascript",         4),  ("cucumber",           5),
    ("bdd",                5), ("testng",             4),  ("junit",              4),
    ("restassured",        5), ("pytest",             4),  ("api testing",        6),
    ("rest api",           5), ("jenkins",            4),  ("github actions",     4),
    ("ci/cd",              4), ("ci cd",              4),  ("aws",                3),
    ("docker",             3), ("salesforce",         4),  ("agile",              3),
    ("scrum",              2), ("jira",               2),  ("senior",             3),
    ("lead",               3), ("staff",              2),  ("principal",          2),
]
JUNIOR_SIGNALS = ["junior", "associate", "entry level", "entry-level", "intern", "0-2 years", "1-2 years"]
TITLE_BOOSTS   = {"sdet":15, "qa automation":12, "test automation":12, "quality engineer":8,
                  "automation engineer":8, "qa lead":10, "qa engineer":7,
                  "software engineer in test":10}
MAX_POSSIBLE   = sum(w for _, w in SKILL_WEIGHTS) + 15

def ats_score_job(job: dict) -> dict:
    text = (job.get("title","") + " " + job.get("snippet","")).lower()
    matched, raw = [], 0
    for keyword, weight in SKILL_WEIGHTS:
        if keyword in text:
            matched.append(keyword); raw += weight
    for phrase, bonus in TITLE_BOOSTS.items():
        if phrase in job.get("title","").lower():
            raw += bonus; break
    if any(s in text for s in JUNIOR_SIGNALS):
        raw = max(0, raw - 20)
    score  = min(100, int((raw / MAX_POSSIBLE) * 100))
    important = ["playwright","selenium","java","python","ci/cd","api testing"]
    gaps   = [k for k in important if k not in matched]
    job["score"]         = score
    job["match_reasons"] = ", ".join(matched[:6])
    job["gaps"]          = ", ".join(gaps)
    job["ai_summary"]    = f"Matched {len(matched)} skills. " + ("Good seniority match." if any(s in text for s in ["senior","lead","staff"]) else "Check seniority level.")
    return job

def ats_score_batch(jobs: list[dict]) -> list[dict]:
    """Score all jobs locally — 100% free, no API calls."""
    for job in jobs:
        ats_score_job(job)
    return jobs


# ─────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────

EXCEL_COLS = [
    "Date Found", "Source", "Company", "Job Title", "ATS Score", "H1B Sponsor",
    "Match Skills", "Gaps", "AI Summary", "Apply URL", "Status", "Notes"
]

def init_excel():
    EXCEL_PATH.parent.mkdir(exist_ok=True)
    if EXCEL_PATH.exists():
        return openpyxl.load_workbook(EXCEL_PATH)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Job Tracker"

    # Header row styling
    header_fill = PatternFill("solid", fgColor="1D3557")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(bottom=Side(style="thin", color="CCCCCC"))

    for col_idx, col_name in enumerate(EXCEL_COLS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Column widths
    widths = [14, 14, 22, 40, 11, 13, 35, 25, 45, 50, 14, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"
    wb.save(EXCEL_PATH)
    return wb


def append_jobs_to_excel(jobs: list[dict]):
    wb = init_excel()
    ws = wb.active
    next_row = ws.max_row + 1

    score_colors = {
        "high":   "C8F7C5",  # green
        "medium": "FFF3CD",  # yellow
        "low":    "F8D7DA",  # red
    }
    status_color = PatternFill("solid", fgColor="E8F4FD")

    for job in jobs:
        score = job.get("score", 0)
        tier  = "high" if score >= 80 else "medium" if score >= 65 else "low"

        row_data = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            job.get("source", "Company Site"),
            job.get("company", ""),
            job.get("title", ""),
            score,
            job.get("h1b_likely", "unknown").upper(),
            job.get("match_reasons", ""),
            job.get("gaps", ""),
            job.get("ai_summary", ""),
            job.get("url", ""),
            "New",
            ""
        ]

        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=next_row, column=col_idx, value=value)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if col_idx == 4:  # ATS score — color by tier
                cell.fill = PatternFill("solid", fgColor=score_colors[tier])
                cell.font = Font(bold=True)
            if col_idx == 10:  # Status
                cell.fill = status_color

        ws.row_dimensions[next_row].height = 36
        next_row += 1

    wb.save(EXCEL_PATH)
    log.info(f"Saved {len(jobs)} jobs to {EXCEL_PATH}")


# ─────────────────────────────────────────────
# EMAIL APPLICATION
# ─────────────────────────────────────────────

EMAIL_TEMPLATE = """\
Subject: Application for {job_title} – {your_name} | Senior SDET / QA Automation Engineer

Dear Hiring Team at {company},

I am writing to express my strong interest in the {job_title} position at {company}.

As a Senior SDET with 4+ years at American Express, I bring:
• Led migration of 1,000+ test scenarios Selenium → Playwright (60% faster execution)
• 70% increase in regression coverage via scalable cross-browser automation framework  
• 500+ automated end-to-end web & API test scenarios (Selenium, RestAssured, Playwright)
• Zero critical defects across multiple production releases as QA sign-off lead
• Deep expertise in Java, Python, Cucumber BDD, Jenkins/GitHub Actions CI/CD, Salesforce

I hold an MS in Computer Science from the University of Central Missouri and am currently
on H1B status — I am actively seeking an employer who can provide visa sponsorship.

I would welcome the opportunity to discuss how my automation expertise can strengthen
{company}'s quality engineering team.

Apply link: {apply_url}

Best regards,
{your_name}
{your_email}
+1 816-768-1825
https://linkedin.com/in/jagadeeshh-b04264176/
"""

def send_application_email(job: dict, to_email: str | None = None):
    """
    Sends application email.
    If to_email is None, sends a digest email to yourself with the job details.
    """
    if not GMAIL_APP_PASS:
        log.warning("GMAIL_APP_PASS not set — skipping email send.")
        return False

    body = EMAIL_TEMPLATE.format(
        job_title   = job["title"],
        company     = job["company"],
        your_name   = YOUR_NAME,
        your_email  = YOUR_EMAIL,
        apply_url   = job.get("url", "See attachment"),
    )

    recipient = to_email or YOUR_EMAIL  # default: email yourself the digest

    msg = MIMEMultipart()
    msg["From"]    = YOUR_EMAIL
    msg["To"]      = recipient
    msg["Subject"] = f"Application: {job['title']} @ {job['company']} | ATS {job.get('score',0)}%"

    msg.attach(MIMEText(body, "plain"))

    # Attach resume — download from RESUME_URL secret or use local file
    resume_path = Path("config/resume.pdf")
    resume_url  = os.getenv("RESUME_URL", "")
    if not resume_path.exists() and resume_url:
        try:
            r = requests.get(resume_url, timeout=20)
            r.raise_for_status()
            resume_path.parent.mkdir(exist_ok=True)
            resume_path.write_bytes(r.content)
            log.info("Resume downloaded from RESUME_URL.")
        except Exception as e:
            log.warning(f"Could not download resume: {e}")
    if resume_path.exists():
        with open(resume_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f'attachment; filename="Sai_Jagadeesh_Hazari_Resume.pdf"')
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(YOUR_EMAIL, GMAIL_APP_PASS)
            smtp.sendmail(YOUR_EMAIL, recipient, msg.as_string())
        log.info(f"Email sent for: {job['title']} @ {job['company']}")
        return True
    except Exception as e:
        log.error(f"Email failed: {e}")
        return False


def send_digest_email(shortlisted: list[dict]):
    """Send yourself a summary digest of all shortlisted jobs."""
    if not GMAIL_APP_PASS or not shortlisted:
        return

    rows = "\n".join(
        f"  [{j['score']}%] {j['title']} @ {j['company']} — H1B:{j.get('h1b_likely','?').upper()}\n"
        f"         {j.get('url','no url')}\n"
        f"         Matches: {j.get('match_reasons','')}"
        for j in shortlisted
    )

    body = f"""H1B Job Hunt — Run Summary {datetime.now().strftime('%Y-%m-%d %H:%M')}
{'='*60}

Found {len(shortlisted)} shortlisted jobs (ATS ≥ {MIN_ATS_SCORE}%):

{rows}

Full tracker saved to: job_tracker.xlsx
"""
    msg = MIMEMultipart()
    msg["From"]    = YOUR_EMAIL
    msg["To"]      = YOUR_EMAIL
    msg["Subject"] = f"[Job Hunt] {len(shortlisted)} new matches — {datetime.now().strftime('%b %d')}"
    msg.attach(MIMEText(body, "plain"))

    # Attach the Excel file
    if EXCEL_PATH.exists():
        with open(EXCEL_PATH, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", 'attachment; filename="job_tracker.xlsx"')
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(YOUR_EMAIL, GMAIL_APP_PASS)
            smtp.sendmail(YOUR_EMAIL, YOUR_EMAIL, msg.as_string())
        log.info("Digest email sent.")
    except Exception as e:
        log.error(f"Digest email failed: {e}")


# ─────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────

def run_pipeline():
    log.info("=" * 55)
    log.info(f"H1B Job Hunt Pipeline — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    log.info("=" * 55)

    seen = load_seen()

    # 1. Scrape
    log.info("STEP 1/5 — Scraping career pages...")
    raw_jobs = [annotate_job(job) for job in scrape_all_companies()]
    log.info(f"  Total raw listings: {len(raw_jobs)}")

    # 2. De-duplicate
    log.info("STEP 2/5 — De-duplicating...")
    new_jobs = []
    for job in raw_jobs:
        jid = job_id(job["title"], job["company"])
        if jid not in seen:
            job["job_id"] = jid
            new_jobs.append(job)
    log.info(f"  New (unseen) listings: {len(new_jobs)}")

    if not new_jobs:
        log.info("No new jobs found. Pipeline done.")
        return

    # 3. ATS Score
    log.info("STEP 3/5 — ATS scoring with Claude...")
    scored_jobs = ats_score_batch(new_jobs)

    # 4. Filter: USA / India + H1B + ATS score threshold
    log.info("STEP 4/5 — Filtering & shortlisting...")
    location_filtered = [j for j in scored_jobs if location_allowed(j)]
    log.info(f"  Location filtered (USA/India): {len(location_filtered)}")
    role_filtered = [j for j in location_filtered if role_allowed(j)]
    log.info(f"  Role filtered (QA/SDET/test automation): {len(role_filtered)}")
    shortlisted = [
        j for j in role_filtered
        if j.get("score", 0) >= MIN_ATS_SCORE
        and j.get("h1b_likely", "unknown") == "yes"
    ]
    shortlisted.sort(key=lambda x: x.get("score", 0), reverse=True)
    log.info(f"  Shortlisted (score≥{MIN_ATS_SCORE}, H1B sponsorship likely): {len(shortlisted)}")

    # 5. Save all scored to Excel
    log.info("STEP 5/5 — Saving to Excel & sending emails...")
    append_jobs_to_excel(scored_jobs)

    # Update seen set with all new jobs
    for j in new_jobs:
        seen.add(j["job_id"])
    save_seen(seen)

    # Send individual applications for top shortlisted
    applied_count = 0
    for job in shortlisted[:MAX_APPLY_PER_RUN]:
        send_application_email(job)
        applied_count += 1
        time.sleep(3)

    # Send digest to yourself
    send_digest_email(shortlisted)

    log.info("─" * 55)
    log.info(f"Done. Scraped: {len(raw_jobs)} | New: {len(new_jobs)} | "
             f"Shortlisted: {len(shortlisted)} | Applied: {applied_count}")
    log.info("─" * 55)


if __name__ == "__main__":
    run_pipeline()
