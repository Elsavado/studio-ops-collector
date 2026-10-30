"""Sources specific to a game & fintech studio: SEC financial statement data, peer company filings, real contracts
filed as SEC exhibits, federal regulations, case law, occupations data and consumer finance complaints."""
import datetime as dt
import os
import re
import time

from .common import get, safe, session
from .public import ACC, FIN, HR, LEGAL, _download, meta_note

SEC_UA = os.environ.get("SEC_USER_AGENT") or "StudioOps research collector admin.ops.salvador@gmail.com"


def _sec():
    s = session()
    s.headers["User-Agent"] = SEC_UA
    return s


def fsds(sink, st, per_run=30):
    """SEC Financial Statement Data Sets: every company's XBRL financial statements, one archive per quarter."""
    s = _sec()
    done = st.setdefault("fsds_done", [])
    today = dt.date.today()
    n = 0
    for y in range(2009, today.year + 1):
        for q in range(1, 5):
            key = f"{y}q{q}"
            if key in done or (y, q) >= (today.year, (today.month - 1) // 3 + 1) or n >= per_run or sink.full():
                continue
            time.sleep(0.5)
            if _download(s, f"https://www.sec.gov/files/dera/data/financial-statement-data-sets/{key}.zip", 2 ** 31,
                         sink, f"{ACC}/SEC Financial Statement Data Sets/{y}/{key} financial statements.zip"):
                done.append(key)
                n += 1
    if n:
        meta_note(sink, f"{ACC}/SEC Financial Statement Data Sets", {
            "source": "SEC DERA Financial Statement Data Sets", "license": "US Government work, public domain",
            "url": "https://www.sec.gov/dera/data/financial-statement-data-sets"})
    return n


PEERS = {  # name: CIK
    "Electronic Arts": 712515, "Take-Two Interactive": 946581, "Roblox": 1315098, "Unity Software": 1810806,
    "Playtika": 1828016, "Zynga": 1439404, "Activision Blizzard": 718877, "Skillz": 1801661, "PlayAGS": 1593548,
    "PayPal": 1633917, "Block (Square)": 1512673, "Affirm": 1820953, "SoFi Technologies": 1818874,
    "Marqeta": 1522540, "Shift4 Payments": 1794669, "Intuit": 896878, "Bill Holdings": 1786352,
    "Toast": 1650164, "Coinbase": 1679788, "Robinhood": 1783879, "Payoneer": 1845815, "Flywire": 1580560,
}
PEER_FORMS = {"10-K": "Annual reports", "10-K405": "Annual reports", "10-Q": "Quarterly reports",
              "DEF 14A": "Proxy statements", "S-1": "IPO prospectus", "S-1/A": "IPO prospectus",
              "424B4": "IPO prospectus", "20-F": "Annual reports"}


def peers(sink, st):
    """Annual, quarterly, proxy and IPO filings of listed game and fintech companies (primary documents)."""
    s = _sec()
    done = set(st.setdefault("peer_done", []))
    n = 0
    for name, cik in PEERS.items():
        if sink.full():
            break
        r = get(s, f"https://data.sec.gov/submissions/CIK{cik:010d}.json")
        if not r:
            continue
        rec = r.json()["filings"]["recent"]
        for form, date, acc, doc in zip(rec["form"], rec["filingDate"], rec["accessionNumber"], rec["primaryDocument"]):
            if form not in PEER_FORMS or not doc or acc in done or sink.full():
                continue
            time.sleep(0.15)
            url = f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc.replace('-', '')}/{doc}"
            ext = doc.rsplit(".", 1)[-1].lower() if "." in doc else "htm"
            if _download(s, url, 300 * 2 ** 20, sink,
                         f"{ACC}/Peer Company Filings/{name}/{PEER_FORMS[form]}/{date} {form}.{ext}"):
                n += 1
            done.add(acc)
        st["peer_done"] = sorted(done)
    return n


CONTRACT_TYPES = {
    "Software licence agreements": '"software license agreement"',
    "Master services agreements": '"master services agreement"',
    "Statements of work": '"statement of work"',
    "SaaS & subscription agreements": '"subscription agreement" software',
    "Game development & publishing agreements": '"video game" "development agreement" OR "publishing agreement" game',
    "Payment processing & merchant agreements": '"payment processing agreement" OR "merchant agreement"',
    "Employment agreements": '"employment agreement"',
    "Consulting & contractor agreements": '"consulting agreement" OR "independent contractor agreement"',
    "Confidentiality & NDAs": '"confidentiality agreement" OR "non-disclosure agreement"',
    "IP assignment & licensing": '"intellectual property assignment" OR "license agreement" patent',
    "Office leases": '"office lease"',
    "Credit & loan agreements": '"credit agreement"',
    "Equity incentive plans": '"equity incentive plan"',
    "Separation agreements": '"separation agreement"',
    "Reseller & distribution agreements": '"reseller agreement" OR "distribution agreement" software',
    "Data processing agreements": '"data processing agreement"',
    "Revenue sharing agreements": '"revenue sharing agreement"',
}


def contracts(sink, st, per_type=60):
    """Real contracts filed as exhibits (EX-10) with SEC filings, found with EDGAR full-text search."""
    s = _sec()
    offs = st.setdefault("contract_from", {})
    seen = set(st.setdefault("contract_seen", []))
    n = 0
    for folder, q in CONTRACT_TYPES.items():
        took, start = 0, offs.get(folder, 0)
        while took < per_type and not sink.full():
            time.sleep(0.3)
            r = get(s, "https://efts.sec.gov/LATEST/search-index", params={"q": q, "from": start})
            hits = (r.json().get("hits", {}).get("hits", []) if r else [])
            if not hits:
                start = 0
                break
            start += len(hits)
            for h in hits:
                src = h.get("_source", {})
                if not str(src.get("file_type", "")).startswith("EX-10") or h["_id"] in seen:
                    continue
                seen.add(h["_id"])
                acc, fname = h["_id"].split(":", 1)
                cik = (src.get("ciks") or ["0"])[0].lstrip("0")
                company = re.sub(r"\s*\(CIK.*", "", (src.get("display_names") or ["Company"])[0])
                url = f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc.replace('-', '')}/{fname}"
                ext = fname.rsplit(".", 1)[-1].lower() if "." in fname else "htm"
                if _download(s, url, 50 * 2 ** 20, sink, f"{LEGAL}/Contracts (SEC exhibits)/{folder}/"
                             f"{src.get('file_date', '')} {safe(company, 60)} {safe(fname.rsplit('.', 1)[0], 40)}.{ext}"):
                    took += 1
                    n += 1
                if took >= per_type:
                    break
        offs[folder] = start
    st["contract_seen"] = sorted(seen)[-200000:]
    return n


ECFR_TITLES = {12: "Banks and Banking", 16: "Commercial Practices (FTC, COPPA)", 17: "Securities and Commodities",
               29: "Labor", 31: "Money and Finance - Treasury (FinCEN)", 37: "Patents, Trademarks and Copyrights"}


def ecfr(sink, st):
    """Current Code of Federal Regulations titles that govern fintech, games, employment and IP (XML)."""
    s = session()
    r = get(s, "https://www.ecfr.gov/api/versioner/v1/titles.json")
    if not r:
        return 0
    latest = {t["number"]: t.get("latest_issue_date") for t in r.json().get("titles", [])}
    n = 0
    for num, name in ECFR_TITLES.items():
        date = latest.get(num)
        if not date or st.get(f"ecfr_{num}") == date or sink.full():
            continue
        if _download(s, f"https://www.ecfr.gov/api/versioner/v1/full/{date}/title-{num}.xml", 2 ** 31, sink,
                     f"{LEGAL}/Federal Regulations (eCFR)/Title {num} {name} ({date}).xml"):
            st[f"ecfr_{num}"] = date
            n += 1
    return n


REPORTERS = {"br": "Bankruptcy Reporter", "del-ch": "Delaware Chancery Reports", "f3d": "Federal Reporter 3d",
             "f-supp-3d": "Federal Supplement 3d", "f-supp-2d": "Federal Supplement 2d", "f4th": "Federal Reporter 4th"}


def caselaw(sink, st, per_run=150):
    """Court decisions from the Caselaw Access Project (public domain): business, bankruptcy and federal courts."""
    s = session()
    vols = st.setdefault("case_vol", {})
    n = 0
    for rep, name in REPORTERS.items():
        v = vols.get(rep, 1)
        while n < per_run and not sink.full():
            if _download(s, f"https://static.case.law/{rep}/{v}.zip", 2 ** 31, sink,
                         f"{LEGAL}/Case Law/{name}/Volume {v:04d}.zip"):
                n += 1
                v += 1
            else:
                if not get(s, f"https://static.case.law/{rep}/{v + 1}.zip", tries=1, stream=True):
                    break  # end of the reporter
                v += 1
        vols[rep] = v
    if n:
        meta_note(sink, f"{LEGAL}/Case Law", {"source": "Caselaw Access Project (Harvard Law School Library)",
                                              "url": "https://static.case.law", "license": "Public domain / CC0"})
    return n


def onet(sink, st):
    """O*NET occupation database: tasks, skills, knowledge, tools and technology for every job."""
    if st.get("onet"):
        return 0
    s = session()
    for ver in ("30_1", "30_0", "29_3", "29_2"):
        if _download(s, f"https://www.onetcenter.org/dl_files/database/db_{ver}_excel.zip", 2 ** 30, sink,
                     f"{HR}/Occupations & Skills (O-NET)/O-NET database {ver.replace('_', '.')} (Excel).zip"):
            st["onet"] = ver
            meta_note(sink, f"{HR}/Occupations & Skills (O-NET)", {
                "source": "O*NET Resource Center", "license": "CC BY 4.0 (US Department of Labor/ETA)"})
            return 1
    return 0


def cfpb(sink, st):
    """CFPB Consumer Complaint Database: every complaint about financial products and payment companies."""
    stamp = st.get("cfpb")
    if stamp and stamp > (dt.date.today() - dt.timedelta(days=30)).isoformat():
        return 0
    s = session()
    if _download(s, "https://files.consumerfinance.gov/ccdb/complaints.csv.zip", 8 * 2 ** 30, sink,
                 f"{FIN}/Consumer Complaints (CFPB)/Consumer complaints {dt.date.today()}.csv.zip"):
        st["cfpb"] = dt.date.today().isoformat()
        return 1
    return 0


SOURCES = {"fsds": fsds, "peers": peers, "contracts": contracts, "ecfr": ecfr, "caselaw": caselaw, "onet": onet,
           "cfpb": cfpb}
