"""Books & knowledge library: open-access books, NIST guides, Wikipedia topic articles, OWASP, CWE/CAPEC."""
import datetime as dt
import html
import re
import time
from urllib.parse import quote

from docx import Document
from docx.shared import Pt

from .common import get, safe, session
from .public import LIB, _download, meta_note

TOPICS = ["game design", "game development", "video games", "game programming", "virtual worlds", "esports",
          "fintech", "financial technology", "payment systems", "digital payments", "banking technology", "blockchain",
          "accounting", "financial accounting", "management accounting", "corporate finance", "auditing", "taxation",
          "contract law", "intellectual property", "employment law", "data protection", "financial regulation",
          "human resource management", "talent management", "organizational behavior", "leadership",
          "software engineering", "agile software development", "project management", "cloud computing", "devops",
          "user experience", "product management", "marketing", "digital marketing", "sales management",
          "customer service", "entrepreneurship", "startups", "business strategy", "information security"]

BOOK_LICENSES = ("cc-by", "cc-by-sa", "cc-by-nd", "cc-by-nc", "cc-by-nc-sa", "cc-by-nc-nd", "cc0", "public-domain")

NIST_SP = """800-12r1 800-30r1 800-34r1 800-37r2 800-39 800-40r4 800-41r1 800-46r2 800-50r1 800-53r5 800-53Ar5
800-53B 800-55v1 800-60v1r1 800-61r2 800-61r3 800-63-3 800-63B 800-82r3 800-83r1 800-84 800-86 800-88r1 800-92
800-94 800-100 800-115 800-116r1 800-123 800-124r2 800-125 800-128 800-137 800-144 800-145 800-146 800-150
800-153 800-160v1r1 800-160v2r1 800-161r1 800-171r3 800-172 800-181r1 800-184 800-190 800-207 800-209 800-210
800-213 800-215 800-218 800-221 800-222""".split()
NIST_OTHER = ["https://nvlpubs.nist.gov/nistpubs/CSWP/NIST.CSWP.29.pdf",
              "https://nvlpubs.nist.gov/nistpubs/ir/2018/NIST.IR.8183.pdf",
              "https://nvlpubs.nist.gov/nistpubs/ir/2021/NIST.IR.8286.pdf",
              "https://nvlpubs.nist.gov/nistpubs/ir/2020/NIST.IR.8259.pdf",
              "https://nvlpubs.nist.gov/nistpubs/ir/2016/NIST.IR.8112.pdf"]

WIKI_CATS = {
    "Game Development": ["Video game development", "Video game design", "Game engines", "Video game programming",
                         "Video game business models", "Free-to-play video games", "Video game monetization",
                         "Esports", "Multiplayer online games", "Mobile games", "Video game publishers",
                         "Video game industry"],
    "Fintech & Payments": ["Financial technology", "Payment systems", "Online payments", "Mobile payments",
                           "Payment service providers", "Digital currencies", "Banking technology", "Open banking",
                           "Credit cards", "Fraud", "Anti-money laundering", "Know your customer"],
    "Accounting & Finance": ["Accounting", "Financial accounting", "Management accounting", "Accounting terminology",
                             "Auditing", "Corporate finance", "Financial statements", "Taxation",
                             "International Financial Reporting Standards", "Generally Accepted Accounting Principles"],
    "Legal & Compliance": ["Contract law", "Intellectual property law", "Copyright law", "Software licenses",
                           "Employment law", "Privacy law", "Information privacy", "Consumer protection",
                           "Financial regulation", "Corporate law"],
    "HR & Workforce": ["Human resource management", "Recruitment", "Employee benefits", "Performance management",
                       "Organizational behavior", "Workplace", "Remote work", "Leadership"],
    "Software Engineering": ["Software engineering", "Software development process", "Agile software development",
                             "Software testing", "Version control systems", "Continuous integration", "DevOps",
                             "Cloud computing", "Software architecture", "Programming languages", "Databases"],
    "Product, Design & UX": ["User interface design", "Usability", "Human-computer interaction", "Product management",
                             "Web design", "Accessibility"],
    "Sales, Marketing & Support": ["Marketing", "Digital marketing", "Sales", "Customer relationship management",
                                   "Customer service", "Advertising", "Search engine optimization"],
    "Business & Strategy": ["Business models", "Startup companies", "Management", "Strategic management",
                            "Project management", "Outsourcing", "Business process"],
}


def nist_pubs(sink, st):
    s = session()
    done = set(st.setdefault("nist_pubs", []))
    urls = [f"https://nvlpubs.nist.gov/nistpubs/SpecialPublications/NIST.SP.{x}.pdf" for x in NIST_SP] + NIST_OTHER
    n = 0
    for u in urls:
        if u in done or sink.full():
            continue
        name = u.rsplit("/", 1)[1]
        time.sleep(1)
        if _download(s, u, 200 * 2 ** 20, sink, f"{LIB}/NIST Security Publications/{name}"):
            n += 1
        st["nist_pubs"].append(u)
    if n:
        meta_note(sink, f"{LIB}/NIST Security Publications", {
            "source": "NIST Special Publications and reports", "url": "https://csrc.nist.gov/publications",
            "license": "US Government work, public domain"})
    return n


def open_books(sink, st, per_topic=8):
    """Open-access books (CC or public domain, with a PDF) found through OpenAlex."""
    s = session()
    seen = st.setdefault("books_seen", [])
    seen_set = set(seen)
    pages = st.setdefault("books_page", {})
    n = 0
    for topic in TOPICS:
        if sink.full():
            break
        page = pages.get(topic, 1)
        time.sleep(1)
        r = get(s, "https://api.openalex.org/works", params={
            "filter": f"title.search:{topic},type:book|book-chapter,open_access.is_oa:true",
            "per-page": 50, "page": page, "sort": "cited_by_count:desc",
            "select": "id,doi,title,publication_year,authorships,best_oa_location,primary_location,type"})
        res = r.json().get("results", []) if r else []
        pages[topic] = page + 1 if res else 1
        took = 0
        for w in res:
            if took >= per_topic or sink.full():
                break
            loc = w.get("best_oa_location") or {}
            lic, pdf = (loc.get("license") or ""), loc.get("pdf_url")
            if w["id"] in seen_set or not pdf or lic not in BOOK_LICENSES:
                continue
            seen.append(w["id"])
            seen_set.add(w["id"])
            title = safe(w.get("title") or w["id"], 110)
            kind = "Books" if w.get("type") == "book" else "Book chapters"
            dest = f"{LIB}/Open Access Books/{safe(topic.title(), 40)}/{kind}"
            time.sleep(1)
            fname = f"{title} ({w.get('publication_year')})"
            if _download(s, pdf, 300 * 2 ** 20, sink, f"{dest}/{fname}.pdf"):
                meta_note(sink, f"{dest}/{fname}", {
                    "title": w.get("title"), "doi": w.get("doi"), "year": w.get("publication_year"),
                    "authors": [a["author"]["display_name"] for a in w.get("authorships", [])][:20],
                    "publisher": ((loc.get("source") or {}).get("display_name")), "license": lic,
                    "pdf": pdf, "found_via": "OpenAlex"})
                took += 1
                n += 1
    return n


def _wiki_doc(title, html_text, url, credit="Wikipedia contributors, CC BY-SA 4.0"):
    d = Document()
    d.styles["Normal"].font.size = Pt(10.5)
    d.core_properties.title = title
    d.core_properties.comments = f"{url} - {credit}"
    d.add_heading(title, 0)
    d.add_paragraph(f"Source: {url} ({credit}). Retrieved {dt.date.today()}.")
    body = re.sub(r"(?s)<(style|script|table|figure|sup)[^>]*>.*?</\1>", "", html_text)
    for tag, inner in re.findall(r"(?s)<(h2|h3|h4|p|li)[^>]*>(.*?)</\1>", body):
        text = html.unescape(re.sub(r"<[^>]+>", "", inner)).strip()
        text = re.sub(r"\[\d+\]", "", text)
        if not text:
            continue
        if tag == "h2":
            d.add_heading(text, 1)
        elif tag in ("h3", "h4"):
            d.add_heading(text, 2)
        elif tag == "li":
            d.add_paragraph(text, style="List Bullet")
        else:
            d.add_paragraph(text)
    return d


def wikipedia(sink, st, per_cat=40):
    s = session()
    seen = set(st.setdefault("wiki_seen", []))
    cont = st.setdefault("wiki_cont", {})
    api = "https://en.wikipedia.org/w/api.php"
    n = 0
    for group, cats in WIKI_CATS.items():
        for cat in cats:
            if sink.full() or cont.get(cat) == "done":
                continue
            params = {"action": "query", "list": "categorymembers", "cmtitle": f"Category:{cat}",
                      "cmtype": "page", "cmlimit": per_cat, "format": "json"}
            if cont.get(cat):
                params["cmcontinue"] = cont[cat]
            r = get(s, api, params=params)
            if not r:
                continue
            d = r.json()
            cont[cat] = d.get("continue", {}).get("cmcontinue", "done")
            for m in d.get("query", {}).get("categorymembers", []):
                t = m["title"]
                if t in seen or sink.full():
                    continue
                time.sleep(0.2)
                pr = get(s, f"https://en.wikipedia.org/api/rest_v1/page/html/{t.replace(' ', '_')}")
                if not pr:
                    continue
                url = f"https://en.wikipedia.org/wiki/{t.replace(' ', '_')}"
                doc = _wiki_doc(t, pr.text, url)
                p = sink.path(f"{LIB}/Wikipedia Knowledge Base/{group}/{safe(cat, 50)}/{safe(t, 100)}.docx")
                doc.save(p)
                sink.wrote(p)
                seen.add(t)
                st["wiki_seen"].append(t)
                n += 1
    if n:
        meta_note(sink, f"{LIB}/Wikipedia Knowledge Base", {
            "source": "English Wikipedia", "license": "CC BY-SA 4.0 (Wikipedia contributors)",
            "note": "Articles converted to DOCX; text only, tables and images left out."})
    return n


KNOWLEDGE = [
    ("owasp_asvs", "OWASP ASVS", "https://github.com/OWASP/ASVS/archive/refs/heads/master.zip",
     "OWASP Application Security Verification Standard.zip", "CC BY-SA 4.0"),
    ("owasp_masvs", "OWASP MASVS (mobile apps)", "https://github.com/OWASP/owasp-masvs/archive/HEAD.zip",
     "OWASP Mobile Application Security Verification Standard.zip", "CC BY-SA 4.0"),
    ("cwe", "MITRE CWE", "https://cwe.mitre.org/data/xml/cwec_latest.xml.zip", "CWE catalog (latest).xml.zip",
     "CWE Terms of Use (free, attribution)"),
    ("spdx", "SPDX licence list", "https://github.com/spdx/license-list-data/archive/HEAD.zip",
     "SPDX license list data.zip", "CC0"),
    ("choosealicense", "Choose a License (licence texts)", "https://github.com/github/choosealicense.com/archive/HEAD.zip",
     "choosealicense.zip", "MIT"),
]


def knowledge(sink, st):
    s = session()
    n = 0
    for key, title, url, name, lic in KNOWLEDGE:
        if sink.full():
            break
        stamp = st.get(f"kb_{key}")
        if stamp and stamp > (dt.date.today() - dt.timedelta(days=90)).isoformat():
            continue  # refresh each quarter
        if _download(s, url, 2 * 2 ** 30, sink, f"{LIB}/Security Knowledge Bases/{title}/{dt.date.today()} {name}"):
            meta_note(sink, f"{LIB}/Security Knowledge Bases/{title}", {"source": title, "url": url, "license": lic})
            st[f"kb_{key}"] = dt.date.today().isoformat()
            n += 1
    return n


VIDEOS = f"{LIB.rsplit('/', 1)[0]}/15 Videos & Talks"
# folder: searches (all words must appear in the item's title, subjects or description)
VIDEO_SEARCHES = {
    "Game Development": ["game development", "game design", "indie game", "gamedev", "game engine", "level design",
                        "game postmortem", "godot", "unity tutorial", "blender"],
    "Fintech & Payments": ["fintech", "payments", "banking technology", "blockchain", "cryptocurrency"],
    "Accounting & Finance": ["accounting", "bookkeeping", "corporate finance", "financial statements"],
    "Legal & Compliance": ["contract law", "intellectual property", "copyright law", "gdpr", "privacy law"],
    "HR & Management": ["human resources", "management", "leadership", "hiring", "remote work"],
    "Software Engineering": ["software engineering", "programming", "devops", "agile", "cloud computing",
                             "computer science lecture"],
    "Product, Design & Marketing": ["user experience", "product management", "marketing", "startup"],
}
VIDEO_OPEN = ("creativecommons.org", "publicdomain")
STOP = {"of", "and", "the", "it", "on"}


def videos(sink, st, per_query=3, item_cap=4 * 2 ** 30):
    """Openly licensed talks, webinars and recordings from the Internet Archive: smallest MP4 + subtitles."""
    s = session()
    seen = st.setdefault("video_seen", [])
    seen_set = set(seen)
    pages = st.setdefault("video_page", {})
    n = 0
    for folder, qs in VIDEO_SEARCHES.items():
        for q in qs:
            if sink.full():
                return n
            page = pages.get(q, 1)
            time.sleep(1)
            r = get(s, "https://archive.org/advancedsearch.php", params={
                "q": f'("{q}") AND mediatype:movies', "fl[]": ["identifier", "title", "licenseurl", "subject",
                                                              "description"],
                "rows": 50, "page": page, "output": "json", "sort[]": "downloads desc"})
            docs = r.json().get("response", {}).get("docs", []) if r else []
            pages[q] = page + 1 if docs else 1
            want = [w for w in q.lower().split() if w not in STOP]
            took = 0
            for doc in docs:
                if took >= per_query or sink.full():
                    break
                ident, lic = doc["identifier"], doc.get("licenseurl") or ""
                if ident in seen_set or not any(x in lic for x in VIDEO_OPEN):
                    continue
                subj = doc.get("subject") or ""
                text = " ".join([str(doc.get("title", "")), " ".join(subj) if isinstance(subj, list) else str(subj),
                                 str(doc.get("description", ""))[:3000]]).lower()
                if not all(w in text for w in want):
                    continue
                seen.append(ident)
                seen_set.add(ident)
                md = get(s, f"https://archive.org/metadata/{ident}")
                files = md.json().get("files", []) if md else []
                vids = {}
                for f in files:  # smallest MP4 per original video
                    nm = f.get("name", "")
                    if nm.lower().endswith(".mp4") and f.get("size"):
                        stem = nm.rsplit(".", 1)[0].replace(".ia", "").rsplit("_512kb", 1)[0]
                        if stem not in vids or int(f["size"]) < int(vids[stem]["size"]):
                            vids[stem] = f
                subs = [f for f in files if f.get("name", "").lower().endswith((".srt", ".vtt"))]
                if not vids:
                    continue
                dest = f"{VIDEOS}/{folder}/{safe(doc.get('title') or ident, 90)}"
                total, got = 0, 0
                for f in sorted(vids.values(), key=lambda f: f["name"])[:25]:
                    total += int(f["size"])
                    if total > item_cap or sink.full():
                        break
                    if _download(s, f"https://archive.org/download/{ident}/{quote(f['name'])}", 2 * 2 ** 30, sink,
                                 f"{dest}/{f['name'].rsplit('/', 1)[-1]}"):
                        got += 1
                for f in subs[:25]:
                    _download(s, f"https://archive.org/download/{ident}/{quote(f['name'])}", 50 * 2 ** 20, sink,
                              f"{dest}/{f['name'].rsplit('/', 1)[-1]}")
                if got:
                    meta_note(sink, dest, {"source": "Internet Archive", "item": f"https://archive.org/details/{ident}",
                                           "title": doc.get("title"), "license": lic, "search": q})
                    took += 1
                    n += 1
    return n


SOURCES = {"videos": videos, "books": open_books, "wikipedia": wikipedia, "knowledge": knowledge}
