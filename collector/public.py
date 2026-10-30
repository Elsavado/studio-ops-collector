"""Openly licensed public data across every field a game & fintech studio works in."""
import datetime as dt
import hashlib
import io
import os
import re
import shutil
import subprocess
import time

from docx import Document
from docx.shared import Pt

from .common import add_sheet, get, new_book, safe, save_book, session

PUB = "Public Data"
GAME = f"{PUB}/01 Game Development"
FIN = f"{PUB}/02 Fintech & Payments"
ACC = f"{PUB}/03 Accounting & Finance"
LEGAL = f"{PUB}/04 Legal & Compliance"
HR = f"{PUB}/05 HR & Workforce"
ENG = f"{PUB}/06 Software Engineering & DevOps"
PROD = f"{PUB}/07 Product, Design & UX"
GTM = f"{PUB}/08 Sales, Marketing & Customer Support"
SECU = f"{PUB}/09 Security & Privacy"
LIB = f"{PUB}/10 Books & Knowledge Library"

OPEN_LICENSES = ("cc-by", "cc0", "cc-zero", "odc", "pddl", "mit", "apache", "bsd", "gpl", "lgpl",
                 "public", "us-pd", "ogl", "etalab", "other-open", "other-pd")

# Topic searches. Each query feeds one topic folder.
# Topic searches for open dataset catalogues. Each query feeds one topic folder.
SEARCHES = {
    GAME: ["video game dataset", "player behavior", "game telemetry", "steam games", "mobile games", "esports",
           "game reviews", "procedural generation", "game analytics", "in-app purchases"],
    FIN: ["payment transactions", "credit card fraud", "fintech", "mobile money", "digital payments", "open banking",
          "cryptocurrency transactions", "loan default", "credit scoring", "anti money laundering"],
    ACC: ["financial statements", "general ledger", "accounting dataset", "audit", "invoices", "xbrl",
          "corporate earnings", "budget", "expense reports", "tax"],
    LEGAL: ["contracts dataset", "legal documents", "court decisions", "privacy policies", "terms of service",
            "regulatory compliance", "gdpr", "intellectual property", "patents software"],
    HR: ["employee attrition", "salaries", "job postings", "hiring", "employee survey", "workforce", "remote work",
         "performance reviews", "diversity workforce", "skills"],
    ENG: ["software repositories", "code review", "bug reports", "continuous integration", "software defects",
          "developer productivity", "devops", "cloud workloads", "api usage", "technical debt"],
    PROD: ["user experience", "usability", "app reviews", "product analytics", "a/b testing", "user interface design",
           "accessibility", "customer journey"],
    GTM: ["customer support tickets", "crm", "sales pipeline", "marketing campaigns", "customer churn",
          "email marketing", "customer reviews", "advertising"],
    SECU: ["phishing", "intrusion detection", "malware", "data breach", "account takeover", "authentication logs"],
}

def title_query(q):
    """Search term that must appear in the title (Zenodo and CKAN both take Lucene syntax)."""
    words = [w for w in q.split() if w.lower() not in ("dataset", "data", "logs")] or q.split()
    return "title:(" + " AND ".join(words) + ")"


def meta_note(sink, folder, info):
    """Record where a folder's data came from (licence, URL) in the run state, not as a file beside it."""
    sink.state.data.setdefault("catalog", {})[folder] = dict(info, retrieved=dt.date.today().isoformat())

# Search term -> subtopic folder (first match wins, so specific keys come first)
SUBTOPICS = [
    ("player", "Player Behaviour & Telemetry"), ("game telemetry", "Player Behaviour & Telemetry"),
    ("game analytics", "Player Behaviour & Telemetry"), ("in-app", "Monetisation & In-App Purchases"),
    ("esports", "Esports"), ("steam", "Game Catalogues & Sales"), ("mobile games", "Game Catalogues & Sales"),
    ("game reviews", "Game Reviews"), ("procedural", "Game Design & Procedural Generation"),
    ("video game", "Game Catalogues & Sales"), ("credit card fraud", "Fraud Detection"), ("fraud", "Fraud Detection"),
    ("money laundering", "AML & KYC"), ("kyc", "AML & KYC"), ("payment", "Payments & Transactions"),
    ("mobile money", "Payments & Transactions"), ("open banking", "Open Banking"), ("crypto", "Crypto & Digital Assets"),
    ("loan", "Lending & Credit"), ("credit scoring", "Lending & Credit"), ("fintech", "Fintech Markets"),
    ("financial statements", "Financial Statements"), ("xbrl", "Financial Statements"),
    ("earnings", "Financial Statements"), ("general ledger", "Bookkeeping & Ledgers"),
    ("accounting", "Bookkeeping & Ledgers"), ("invoice", "Invoices & Billing"), ("expense", "Expenses"),
    ("tax", "Tax"), ("budget", "Budgeting & Planning"), ("contracts", "Contracts"), ("terms of service", "Policies & Terms"),
    ("privacy polic", "Policies & Terms"), ("court", "Case Law & Court Decisions"), ("patent", "Intellectual Property"),
    ("intellectual property", "Intellectual Property"), ("regulatory", "Regulation & Compliance"),
    ("legal", "Legal Documents"), ("salar", "Compensation"), ("job posting", "Recruiting & Job Postings"),
    ("hiring", "Recruiting & Job Postings"), ("employee survey", "Employee Experience"),
    ("remote work", "Employee Experience"), ("performance review", "Performance Management"),
    ("diversity", "Diversity & Inclusion"), ("skills", "Skills & Occupations"), ("workforce", "Workforce Analytics"),
    ("code review", "Code Review"), ("bug", "Bugs & Defects"), ("defect", "Bugs & Defects"),
    ("continuous integration", "CI/CD & DevOps"), ("devops", "CI/CD & DevOps"), ("software repositor", "Code Repositories"),
    ("developer productivity", "Developer Productivity"), ("technical debt", "Code Quality"),
    ("api usage", "APIs"), ("user experience", "UX Research"), ("usability", "UX Research"),
    ("app reviews", "App Reviews"), ("product analytics", "Product Analytics"), ("a/b test", "Experimentation"),
    ("user interface", "UI Design"), ("accessibility", "Accessibility"), ("customer journey", "Product Analytics"),
    ("support ticket", "Customer Support"), ("sales pipeline", "Sales"), ("marketing", "Marketing"),
    ("advertising", "Marketing"), ("customer reviews", "Customer Feedback"), ("authentication", "Identity & Access"),
    ("account takeover", "Identity & Access"),
    ("intrusion video", "Video Analytics"), ("anomaly detection video", "Video Analytics"),
    ("person detection", "Video Analytics"), ("weapon", "Video Analytics"), ("face mask", "Video Analytics"),
    ("crowd", "Video Analytics"), ("license plate", "Video Analytics"),
    ("software vulnerab", "Vulnerabilities"), ("intrusion", "Intrusion Detection"),
    ("network traffic", "Network Traffic"), ("botnet", "Botnets & DDoS"), ("ddos", "Botnets & DDoS"),
    ("phishing", "Phishing & Fraud"), ("fraud", "Phishing & Fraud"), ("malware", "Malware & Ransomware"),
    ("ransomware", "Malware & Ransomware"), ("honeypot", "Honeypots"), ("vulnerab", "Vulnerabilities"),
    ("cve", "Vulnerabilities"), ("threat intelligence", "Threat Intelligence Feeds"),
    ("cyber threat", "Threat Intelligence Feeds"), ("security operations", "SOC Alerts & SIEM"),
    ("siem", "SOC Alerts & SIEM"), ("security logs", "SOC Alerts & SIEM"), ("insider", "Insider Threat"),
    ("breach", "Data Breaches"), ("privacy", "Privacy & Data Protection"), ("gdpr", "Privacy & Data Protection"),
    ("cyber", "Cyber Incidents"), ("compliance", "Compliance & Audit"), ("audit", "Compliance & Audit"),
    ("risk", "Risk Management"), ("policy", "Security Policy"), ("information security", "Security Policy"),
    ("311", "Service Requests (311)"), ("service requests", "Service Requests (311)"),
    ("event log", "Process Event Logs"), ("service management", "Incident & Service Tickets"),
    ("help desk", "Incident & Service Tickets"), ("helpdesk", "Incident & Service Tickets"),
    ("service desk", "Incident & Service Tickets"), ("tickets", "Incident & Service Tickets"),
    ("servicenow", "Incident & Service Tickets"), ("itsm", "Incident & Service Tickets"),
    ("work orders", "Maintenance & Work Orders"), ("maintenance", "Maintenance & Work Orders"),
    ("asset", "IT Asset Management"), ("surveillance", "Video Surveillance"), ("cctv", "Video Surveillance"),
    ("camera", "Video Surveillance"), ("access control", "Access Control"), ("security guard", "Guarding & Patrol"),
    ("alarm", "Alarms & Fire Safety"), ("fire", "Alarms & Fire Safety"), ("smoke", "Alarms & Fire Safety"),
    ("crime", "Crime & Calls for Service"), ("calls for service", "Crime & Calls for Service"),
    ("burglary", "Crime & Calls for Service"), ("dispatch", "Crime & Calls for Service"),
    ("security incidents", "Crime & Calls for Service"), ("hotel", "Hotels & Lodging"),
    ("lodging", "Hotels & Lodging"), ("short term rental", "Hotels & Lodging"), ("airbnb", "Hotels & Lodging"),
    ("hospitality", "Hotels & Lodging"), ("restaurant", "Hotels & Lodging"), ("cruise", "Cruise & Vessels"),
    ("vessel", "Cruise & Vessels"), ("ais", "Cruise & Vessels"), ("ship", "Cruise & Vessels"),
    ("port", "Ports & Marinas"), ("marina", "Ports & Marinas"), ("tourism", "Tourism"), ("visitor", "Tourism"),
    ("server logs", "System & Server Logs"), ("system log", "System & Server Logs"),
    ("log anomaly", "System & Server Logs"), ("web server", "System & Server Logs"),
    ("cloud", "Cloud Workloads & Costs"), ("kubernetes", "Cloud Workloads & Costs"),
    ("cpu", "Infrastructure Metrics"), ("telemetry", "Infrastructure Metrics"), ("iot", "Infrastructure Metrics"),
    ("data center", "Data Centers"), ("network", "Network Monitoring & Outages"),
    ("outage", "Network Monitoring & Outages"), ("uptime", "Network Monitoring & Outages"),
    ("latency", "Network Monitoring & Outages"), ("broadband", "Network Monitoring & Outages"),
    ("internet", "Network Monitoring & Outages"), ("salar", "Salaries & Payroll"), ("wages", "Salaries & Payroll"),
    ("payroll", "Salaries & Payroll"), ("technician", "Salaries & Payroll"), ("layoffs", "Workforce"),
    ("attrition", "Workforce"), ("hr analytics", "Workforce"), ("occupational", "Workforce"),
    ("security companies", "Security Industry"), ("security services", "Security Industry"),
    ("contracts", "Contracts & Procurement"), ("procurement", "Contracts & Procurement"),
    ("purchase orders", "Contracts & Procurement"), ("vendor", "Contracts & Procurement"),
    ("expenditures", "Spending"), ("spending", "Spending"), ("facility", "Facility Management"),
    ("saas", "Sales & Customers"), ("crm", "Sales & Customers"), ("churn", "Sales & Customers"),
    ("b2b", "Sales & Customers"), ("invoices", "Accounting"), ("accounting", "Accounting"),
    ("stack overflow", "Developer Surveys"), ("developer survey", "Developer Surveys"),
    ("github", "Open Source & Repositories"), ("open source", "Open Source & Repositories"),
    ("software projects", "Open Source & Repositories"), ("programming", "Programming Languages"),
    ("bugs", "Software Quality"), ("app reviews", "Software Quality"), ("devops", "DevOps & Containers"),
    ("docker", "DevOps & Containers"), ("linux", "Operating Systems"), ("tech stack", "APIs & Tech Stacks"),
    ("api", "APIs & Tech Stacks"),
]


def subtopic(q):
    ql = q.lower()
    return next((name for key, name in SUBTOPICS if key in ql), q.title())


# ---------- fixed sources ----------

def _download(s, url, cap, sink, rel):
    """Stream url to the sink path; drop it if it is bigger than cap. Returns True when saved."""
    p = sink.path(rel)
    try:
        with s.get(url, stream=True, timeout=180) as r:
            if r.status_code != 200 or int(r.headers.get("Content-Length") or 0) > cap:
                return False
            n = 0
            with open(p, "wb") as f:
                for chunk in r.iter_content(1 << 20):
                    n += len(chunk)
                    if n > cap:
                        break
                    f.write(chunk)
        if n > cap or n == 0:
            os.remove(p)
            return False
        if p.lower().endswith((".htm", ".html")):
            p = html_to_pdf(p)
        sink.wrote(p)
        return True
    except Exception:
        if os.path.exists(p):
            os.remove(p)
        return False


def html_to_pdf(p):
    """Web-page documents (SEC filings, contract exhibits) are kept as PDF; the HTML is removed when it converts."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return p
    outdir = os.path.dirname(p)
    try:
        subprocess.run([soffice, "--headless", "--norestore", "--convert-to", "pdf:writer_web_pdf_Export",
                        "--outdir", outdir, p], capture_output=True, timeout=300)
    except (subprocess.TimeoutExpired, OSError):
        return p
    pdf = os.path.splitext(p)[0] + ".pdf"
    if os.path.exists(pdf) and os.path.getsize(pdf) > 0:
        os.remove(p)
        return pdf
    return p


def _open(lic):
    lic = (lic or "").lower()
    return any(lic.startswith(x) or x in lic for x in OPEN_LICENSES) and "closed" not in lic \
        and "nc" not in lic.split("-")


def zenodo(sink, st, per_query=6, cap=2 * 2 ** 30):
    s = session()
    seen = st.setdefault("zenodo_seen", [])
    seen_set = set(seen)
    pages = st.setdefault("zenodo_page", {})
    n = 0
    for folder, qs in SEARCHES.items():
        for q in qs:
            if sink.full():
                return n
            page = pages.get(q, 1)
            time.sleep(2.5)
            r = get(s, "https://zenodo.org/api/records", params={
                "q": title_query(q), "size": 25, "page": page, "sort": "mostviewed", "access_status": "open"})
            hits = r.json().get("hits", {}).get("hits", []) if r else []
            pages[q] = page + 1 if hits else 1
            took = 0
            for h in hits:
                if took >= per_query or sink.full():
                    break
                rid = str(h["id"])
                md = h.get("metadata", {})
                lic = (md.get("license") or {}).get("id", "")
                if rid in seen_set or not _open(lic):
                    continue
                seen.append(rid)
                seen_set.add(rid)
                files = [f for f in h.get("files", []) if f["key"].rsplit(".", 1)[-1].lower() in KEEP_EXT
                         and f.get("size", 0) <= cap]
                if not files:
                    continue
                dest = f"{folder}/{subtopic(q)}/{safe(md.get('title', rid), 90)}"
                got = 0
                for f in files[:12]:
                    time.sleep(1)
                    if _download(s, f["links"]["self"], cap, sink, f"{dest}/{f['key']}"):
                        got += 1
                if got:
                    meta_note(sink, dest, {"source": "Zenodo", "record": h.get("links", {}).get("self_html"),
                                           "doi": h.get("doi"), "title": md.get("title"),
                                           "license": lic, "search": q,
                                           "description": (md.get("description") or "")[:4000]})
                    took += 1
                    n += 1
    return n


# Searches suited to government open-data portals; every word must be in the dataset name.
# Searches suited to government open-data portals; every word must be in the dataset name.
PORTAL_SEARCHES = {
    GAME: ["video game", "esports", "gaming"],
    FIN: ["consumer complaints", "payments", "money transmitter", "fraud", "credit card", "banking"],
    ACC: ["expenditures", "general ledger", "vendor payments", "budget", "revenue", "audit", "financial statements"],
    LEGAL: ["contracts", "lawsuits", "settlements", "regulations", "licenses", "court cases"],
    HR: ["employee salaries", "payroll", "job postings", "workforce", "employee demographics", "training"],
    ENG: ["it projects", "software", "technology spending", "data center", "it assets"],
    PROD: ["website analytics", "digital services", "user satisfaction"],
    GTM: ["customer service", "311 service requests", "call center", "customer satisfaction"],
    SECU: ["data breach", "cybersecurity", "privacy"],
}

STOP = {"for", "of", "and", "the", "it"}


def _open_portal(lic):
    lic = (lic or "").lower()
    if not lic:
        return True  # government portals without a stated licence publish public records
    if "noncommercial" in lic or "non-commercial" in lic or "no derivatives" in lic:
        return False
    return any(x in lic for x in ("public domain", "creative commons", "open data commons", "odc", "open government",
                                  "cc0", "attribution", "open database"))


def opendata(sink, st, per_query=5, cap=1 * 2 ** 30):
    """US city / state / federal open-data portals through the Socrata Discovery API (CSV exports)."""
    s = session()
    seen = st.setdefault("socrata_seen", [])
    seen_set = set(seen)
    offs = st.setdefault("socrata_offset", {})
    n = 0
    for folder, qs in PORTAL_SEARCHES.items():
        for q in qs:
            if sink.full():
                return n
            off = offs.get(q, 0)
            time.sleep(1)
            r = get(s, "https://api.us.socrata.com/api/catalog/v1",
                    params={"q": q, "only": "dataset", "limit": 20, "offset": off, "order": "relevance"})
            res = r.json().get("results", []) if r else []
            offs[q] = off + 20 if res else 0
            want = [w for w in q.lower().split() if w not in STOP] or q.lower().split()
            took = 0
            for x in res:
                if took >= per_query or sink.full():
                    break
                rs, md = x["resource"], x["metadata"]
                key = f"{md['domain']}/{rs['id']}"
                name = rs.get("name") or rs["id"]
                if key in seen_set or not _open_portal(md.get("license")):
                    continue
                words = set(re.findall(r"[a-z0-9]+", name.lower()))
                if not all(w in words or w + "s" in words for w in want):
                    continue
                seen.append(key)
                seen_set.add(key)
                dest = f"{folder}/{subtopic(q)}/{safe(name, 80)} ({md['domain']})"
                url = f"https://{md['domain']}/api/views/{rs['id']}/rows.csv?accessType=DOWNLOAD"
                if _download(s, url, cap, sink, f"{dest}/{safe(name, 80)}.csv"):
                    meta_note(sink, dest, {"source": "Socrata open-data portal", "portal": md["domain"],
                                           "dataset": x.get("permalink") or x.get("link"),
                                           "publisher": (x.get("classification") or {}).get("domain_category"),
                                           "license": md.get("license") or "not stated (government open data)",
                                           "updated": rs.get("data_updated_at"), "search": q,
                                           "columns": rs.get("columns_name", [])[:200],
                                           "description": (rs.get("description") or "")[:4000]})
                    took += 1
                    n += 1
    return n


SOURCES = {"zenodo": zenodo, "opendata": opendata}
