"""Named datasets: specific Kaggle datasets and competition data, UCI datasets, OpenDota esports matches, Riot Data
Dragon, and topic searches of the Hugging Face Hub."""
import datetime as dt
import os
import time

from .common import get, safe, session
from .public import ACC, ENG, FIN, GAME, GTM, HR, LEGAL, PROD, SECU, _download, meta_note, subtopic

KAGGLE_DATASETS = {  # folder: Kaggle dataset refs
    f"{FIN}/Fraud Detection": ["ealaxi/paysim1"],
    f"{FIN}/Lending & Credit": ["wordsforthewise/lending-club"],
}
KAGGLE_COMPETITIONS = {f"{FIN}/Fraud Detection": ["ieee-fraud-detection"]}


def kaggle_named(sink, st):
    """Specific Kaggle datasets and competition data (a competition needs its rules accepted on the Kaggle account)."""
    token = os.environ.get("KAGGLE_API_TOKEN")
    if not token:
        print("kaggle_named: KAGGLE_API_TOKEN not set, skipped")
        return 0
    s = session()
    s.headers["Authorization"] = f"Bearer {token}"
    done = st.setdefault("kaggle_named", [])
    n = 0
    jobs = [(f, f"datasets/download/{r}", r) for f, refs in KAGGLE_DATASETS.items() for r in refs] + \
           [(f, f"competitions/data/download-all/{c}", c) for f, cs in KAGGLE_COMPETITIONS.items() for c in cs]
    for folder, path, ref in jobs:
        if ref in done or sink.full():
            continue
        if _download(s, f"https://www.kaggle.com/api/v1/{path}", 8 * 2 ** 30, sink,
                     f"{folder}/{safe(ref.replace('/', ' - '), 80)}.zip"):
            done.append(ref)
            n += 1
        else:
            print(f"kaggle_named: {ref} not downloaded (competition rules not accepted, or no access)")
    return n


UCI = {  # UCI Machine Learning Repository id: (name, folder)
    350: ("Default of credit card clients", f"{FIN}/Lending & Credit"),
    144: ("Statlog German credit data", f"{FIN}/Lending & Credit"),
    222: ("Bank marketing", f"{GTM}/Marketing"),
    352: ("Online retail", f"{GTM}/Sales"),
    365: ("Polish companies bankruptcy", f"{ACC}/Financial Statements"),
    502: ("Online retail II", f"{GTM}/Sales"),
    468: ("Online shoppers purchasing intention", f"{PROD}/Product Analytics"),
}


def uci(sink, st):
    s = session()
    done = st.setdefault("uci", [])
    n = 0
    for uid, (name, folder) in UCI.items():
        if uid in done or sink.full():
            continue
        slug = name.lower().replace(" ", "+")
        for url in (f"https://archive.ics.uci.edu/static/public/{uid}/{slug}.zip",
                    f"https://archive.ics.uci.edu/static/public/{uid}/data.zip"):
            if _download(s, url, 2 ** 30, sink, f"{folder}/{name} (UCI).zip"):
                done.append(uid)
                n += 1
                break
    return n


def opendota(sink, st, matches_per_run=1500):
    """Dota 2 esports data from OpenDota: pro match details (picks, bans, players, gold/xp curves), heroes, items,
    teams and leagues. Free API, 60 calls a minute; walks back through pro matches run by run."""
    s = session()
    api = "https://api.opendota.com/api"
    base = f"{GAME}/Esports Data/Dota 2 (OpenDota)"
    stamp = st.get("od_constants")
    if not stamp or stamp < (dt.date.today() - dt.timedelta(days=30)).isoformat():
        for name in ("heroes", "items", "abilities", "game_mode", "lobby_type", "patch", "region"):
            r = get(s, f"{api}/constants/{name}")
            if r:
                sink.write_bytes(f"{base}/Reference/{name}.json", r.content)
        for name in ("teams", "leagues", "proPlayers"):
            r = get(s, f"{api}/{name}")
            if r:
                sink.write_bytes(f"{base}/Reference/{name} {dt.date.today()}.json", r.content)
        st["od_constants"] = dt.date.today().isoformat()
    last = st.get("od_less_than")
    got = 0
    while got < matches_per_run and not sink.full():
        r = get(s, f"{api}/proMatches", params={"less_than_match_id": last} if last else {})
        lst = r.json() if r else []
        if not lst:
            break
        for m in lst:
            if got >= matches_per_run or sink.full():
                break
            time.sleep(1.05)
            d = get(s, f"{api}/matches/{m['match_id']}")
            if d:
                month = dt.datetime.fromtimestamp(m.get("start_time") or 0, dt.timezone.utc).strftime("%Y-%m")
                sink.write_bytes(f"{base}/Pro matches/{month}/{m['match_id']}.json", d.content)
                got += 1
            last = m["match_id"]
        st["od_less_than"] = last
    if got:
        meta_note(sink, base, {"source": "OpenDota API", "url": "https://www.opendota.com", "license": "CC BY-SA 4.0"})
    return got


def datadragon(sink, st):
    """Riot Games Data Dragon: League of Legends game data and art for the current patch (free static files)."""
    s = session()
    r = get(s, "https://ddragon.leagueoflegends.com/api/versions.json")
    ver = r.json()[0] if r else None
    if not ver or st.get("ddragon") == ver:
        return 0
    url = f"https://ddragon.leagueoflegends.com/cdn/dragontail-{ver}.tgz"
    size = int(s.head(url, timeout=60).headers.get("Content-Length") or 0)
    rel = f"{GAME}/Esports Data/League of Legends (Riot Data Dragon)/dragontail-{ver}.tgz"
    ok = sink.copyurl(url, rel, size) if size > 3 * 2 ** 30 else _download(s, url, 3 * 2 ** 30, sink, rel)
    if ok:
        st["ddragon"] = ver
        meta_note(sink, rel.rsplit("/", 1)[0], {"source": "Riot Games Data Dragon", "version": ver,
                                                "license": "Riot Games Legal Jibber Jabber (non-commercial use)"})
    return int(ok)


HF_SEARCHES = {  # folder: Hugging Face dataset searches
    GAME: ["video game", "steam", "game reviews", "esports", "dota", "league of legends", "minecraft", "chess games"],
    FIN: ["fraud", "credit card", "transactions", "fintech", "financial news", "loan", "banking"],
    ACC: ["financial statements", "sec filings", "10-k", "accounting", "invoices", "financial reports"],
    LEGAL: ["legal", "contracts", "case law", "privacy policy", "terms of service", "patents"],
    HR: ["job postings", "resume", "salaries", "employee", "hr"],
    ENG: ["code", "github issues", "commit messages", "code review", "stack overflow", "software engineering"],
    PROD: ["app reviews", "user feedback", "ux", "product reviews"],
    GTM: ["customer support", "sales emails", "marketing", "customer service dialogue", "crm"],
    SECU: ["phishing", "malware", "cybersecurity", "intrusion detection"],
}
HF_OPEN = {"apache-2.0", "mit", "cc-by-4.0", "cc-by-sa-4.0", "cc0-1.0", "cc-by-3.0", "cc-by-sa-3.0", "odc-by", "odbl",
           "pddl", "bsd-3-clause", "bsd-2-clause", "cdla-permissive-2.0", "cdla-permissive-1.0", "openrail",
           "unlicense", "gpl-3.0", "cc-by-nc-4.0", "cc-by-nc-sa-4.0", "cc-by-2.0", "afl-3.0"}


def huggingface(sink, st, per_query=3, dataset_cap=5 * 2 ** 30):
    """Openly licensed Hugging Face datasets for every studio field (resumes mid-dataset on the next run)."""
    s = session()
    if os.environ.get("HF_TOKEN"):
        s.headers["Authorization"] = f"Bearer {os.environ['HF_TOKEN']}"
    api = "https://huggingface.co/api/datasets"
    seen = st.setdefault("hf_seen", [])
    seen_set = set(seen)
    partial = st.setdefault("hf_partial", {})
    n = 0

    def fetch(rid, folder, start):
        r = get(s, f"{api}/{rid}/tree/main", params={"recursive": "true"})
        files = sorted((f for f in (r.json() if r else []) if f.get("type") == "file"
                        and not f["path"].startswith(".") and not f["path"].endswith(".gitattributes")),
                       key=lambda f: f["path"])
        if not files or sum(int(f.get("size") or 0) for f in files) > dataset_cap:
            partial.pop(rid, None)
            return 0
        dest = f"{folder}/{safe(rid.replace('/', ' - '), 100)}"
        for i in range(start, len(files)):
            if sink.full():
                partial[rid] = [folder, i]
                return 0
            _download(s, f"https://huggingface.co/datasets/{rid}/resolve/main/{files[i]['path']}", 3 * 2 ** 30, sink,
                      f"{dest}/{files[i]['path']}")
        partial.pop(rid, None)
        return 1

    for rid, (folder, start) in list(partial.items()):
        n += fetch(rid, folder, start)
    for top, queries in HF_SEARCHES.items():
        for q in queries:
            if sink.full():
                return n
            r = get(s, api, params={"search": q, "sort": "downloads", "limit": 50, "full": "true"})
            took = 0
            for d in (r.json() if r else []):
                if took >= per_query or sink.full():
                    break
                rid = d["id"]
                lic = (d.get("cardData") or {}).get("license")
                lic = (lic[0] if isinstance(lic, list) and lic else lic) or next(
                    (t.split(":", 1)[1] for t in d.get("tags", []) if t.startswith("license:")), "")
                if rid in seen_set or d.get("gated") or d.get("private") or str(lic).lower() not in HF_OPEN:
                    continue
                seen.append(rid)
                seen_set.add(rid)
                got = fetch(rid, f"{top}/{subtopic(q)}", 0)
                if got:
                    sink.state.data.setdefault("catalog", {})[f"{top}/{rid}"] = {
                        "source": f"https://huggingface.co/datasets/{rid}", "license": lic, "search": q}
                took += got
                n += got
    return n


SOURCES = {"kaggle_named": kaggle_named, "uci": uci, "opendota": opendota, "datadragon": datadragon,
           "huggingface": huggingface}
