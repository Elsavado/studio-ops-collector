"""Kaggle datasets with open licences. Listing is anonymous; downloads need KAGGLE_API_TOKEN
(or the older KAGGLE_USERNAME / KAGGLE_KEY pair)."""
import os
import time

from .common import get, safe, session
from .public import ACC, ENG, FIN, GAME, GTM, HR, LEGAL, PROD, SECU, _download, meta_note, subtopic

SEARCHES = {
    GAME: ["video game sales", "steam games", "player behavior", "game reviews", "mobile games", "esports",
           "game churn", "in-app purchases", "twitch", "game telemetry", "board games", "video game"],
    FIN: ["credit card fraud", "payment transactions", "fintech", "bank transactions", "loan default", "credit risk",
          "money laundering", "cryptocurrency", "stock market", "mobile money", "insurance claims"],
    ACC: ["financial statements", "accounting", "invoices", "general ledger", "audit", "company financials",
          "budget", "expenses", "tax"],
    LEGAL: ["legal documents", "contracts", "court cases", "privacy policies", "patents", "terms of service"],
    HR: ["employee attrition", "hr analytics", "salaries", "data science salaries", "job postings", "resume",
         "employee performance", "layoffs", "glassdoor", "stack overflow survey"],
    ENG: ["github repositories", "software bugs", "code", "programming languages", "developer survey", "devops",
          "cloud", "server logs", "api", "software projects"],
    PROD: ["app reviews", "google play store", "app store", "user behavior", "a/b testing", "ux"],
    GTM: ["customer support tickets", "customer churn", "crm", "marketing campaign", "sales", "customer reviews",
          "ecommerce", "email marketing"],
    SECU: ["phishing", "intrusion detection", "malware", "cybersecurity"],
}
OPEN = ("cc0", "cc-by", "cc by", "attribution", "odbl", "open database", "odc", "public domain", "mit",
        "apache", "gpl", "lgpl", "bsd", "community data license", "cdla", "world bank", "us government",
        "open government")


def _ok(lic):
    lic = (lic or "").lower()
    return any(x in lic for x in OPEN) and "noncommercial" not in lic and "nc-" not in lic


def kaggle(sink, st, per_query=3, cap=3 * 2 ** 30):
    token = os.environ.get("KAGGLE_API_TOKEN")
    user, key = os.environ.get("KAGGLE_USERNAME"), os.environ.get("KAGGLE_KEY")
    if not (token or (user and key)):
        print("kaggle: no Kaggle credentials, skipped")
        return 0
    s = session()
    if token:
        s.headers["Authorization"] = f"Bearer {token}"
    else:
        s.auth = (user, key)
    seen = st.setdefault("kaggle_seen", [])
    seen_set = set(seen)
    pages = st.setdefault("kaggle_page", {})
    n = 0
    for folder, qs in SEARCHES.items():
        for q in qs:
            if sink.full():
                return n
            page = pages.get(q, 1)
            time.sleep(1)
            r = get(s, "https://www.kaggle.com/api/v1/datasets/list",
                    params={"search": q, "sortBy": "votes", "page": page})
            res = r.json() if r else []
            pages[q] = page + 1 if res else 1
            took = 0
            for d in res:
                if took >= per_query or sink.full():
                    break
                ref = d.get("ref", "")
                if not ref or ref in seen_set or not _ok(d.get("licenseName")) or d.get("totalBytes", 0) > cap:
                    continue
                seen.append(ref)
                seen_set.add(ref)
                dest = f"{folder}/{subtopic(q)}/{safe(d.get('title') or ref, 80)}"
                time.sleep(1)
                if _download(s, f"https://www.kaggle.com/api/v1/datasets/download/{ref}", cap, sink,
                             f"{dest}/{safe(ref.replace('/', ' - '), 100)}.zip"):
                    meta_note(sink, dest, {"source": "Kaggle", "dataset": f"https://www.kaggle.com/datasets/{ref}",
                                           "title": d.get("title"), "subtitle": d.get("subtitle"),
                                           "license": d.get("licenseName"), "updated": d.get("lastUpdated"),
                                           "votes": d.get("voteCount"), "downloads": d.get("downloadCount"),
                                           "search": q})
                    took += 1
                    n += 1
    return n


SOURCES = {"kaggle": kaggle}
