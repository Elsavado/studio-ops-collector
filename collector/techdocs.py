"""IT & developer tools: product lifecycles, Q&A archives, official documentation, tool catalogues."""
import datetime as dt
import os
import re
import time
from urllib.parse import quote

from .common import add_sheet, get, new_book, safe, save_book, session
from .library import _wiki_doc
from .public import PUB, _download, meta_note

TECH = f"{PUB}/06 Software Engineering & DevOps"
QA = f"{PUB}/14 Q&A Archives"

SE_SITES = ["gamedev.stackexchange.com", "stackoverflow.com", "softwareengineering.stackexchange.com",
            "codereview.stackexchange.com", "money.stackexchange.com", "quant.stackexchange.com",
            "law.stackexchange.com", "workplace.stackexchange.com", "pm.stackexchange.com", "ux.stackexchange.com",
            "economics.stackexchange.com", "security.stackexchange.com", "crypto.stackexchange.com",
            "datascience.stackexchange.com", "dba.stackexchange.com", "devops.stackexchange.com",
            "serverfault.com", "superuser.com", "unix.stackexchange.com", "webapps.stackexchange.com",
            "freelancing.stackexchange.com", "startups.stackexchange.com", "patents.stackexchange.com",
            "opensource.stackexchange.com", "bitcoin.stackexchange.com", "ethereum.stackexchange.com",
            "gamedev.meta.stackexchange.com", "boardgames.stackexchange.com", "gaming.stackexchange.com"]

# (folder, title, url, file name, licence)
DOCS = [
    (f"{PUB}/01 Game Development/Engine Documentation", "Godot Engine documentation", "https://github.com/godotengine/godot-docs/archive/HEAD.zip",
     "godot-docs.zip", "CC BY 3.0"),
    (f"{PUB}/01 Game Development/Engine Documentation", "Bevy engine (source and examples)", "https://github.com/bevyengine/bevy/archive/HEAD.zip",
     "bevy.zip", "MIT / Apache-2.0"),
    (f"{PUB}/01 Game Development/Engine Documentation", "Phaser examples", "https://github.com/phaserjs/examples/archive/HEAD.zip",
     "phaser-examples.zip", "MIT"),
    (f"{PUB}/01 Game Development/Engine Documentation", "Defold manuals", "https://github.com/defold/doc/archive/HEAD.zip", "defold-doc.zip",
     "CC BY 4.0"),
    (f"{PUB}/01 Game Development/Engine Documentation", "raylib", "https://github.com/raysan5/raylib/archive/HEAD.zip", "raylib.zip", "zlib"),
    ("Web & App Documentation", "MDN web docs", "https://github.com/mdn/content/archive/HEAD.zip", "mdn-content.zip",
     "CC BY-SA 2.5"),
    ("Web & App Documentation", "React documentation", "https://github.com/reactjs/react.dev/archive/HEAD.zip",
     "react.dev.zip", "CC BY 4.0"),
    ("Web & App Documentation", "TypeScript handbook", "https://github.com/microsoft/TypeScript-Website/archive/HEAD.zip",
     "TypeScript-Website.zip", "MIT / CC BY 4.0"),
    ("Web & App Documentation", "Kotlin documentation", "https://github.com/JetBrains/kotlin-web-site/archive/HEAD.zip",
     "kotlin-web-site.zip", "Apache-2.0"),
    ("Cloud & DevOps Documentation", "Kubernetes documentation", "https://github.com/kubernetes/website/archive/HEAD.zip",
     "kubernetes-website.zip", "CC BY 4.0"),
    ("Cloud & DevOps Documentation", "Docker documentation", "https://github.com/docker/docs/archive/HEAD.zip",
     "docker-docs.zip", "Apache 2.0"),
    ("Cloud & DevOps Documentation", "PostgreSQL manual", "https://www.postgresql.org/files/documentation/pdf/17/postgresql-17-A4.pdf",
     "PostgreSQL 17 manual (A4).pdf", "PostgreSQL Licence"),
    ("Cloud & DevOps Documentation", "Pro Git book", "https://github.com/progit/progit2/archive/HEAD.zip", "progit2.zip",
     "CC BY-NC-SA 3.0"),
    ("Engineering Handbooks", "GitLab handbook", "https://gitlab.com/gitlab-com/content-sites/handbook/-/archive/main/handbook-main.zip",
     "gitlab-handbook.zip", "CC BY-SA 4.0"),
    ("Engineering Handbooks", "Microsoft Engineering Playbook",
     "https://github.com/microsoft/code-with-engineering-playbook/archive/HEAD.zip", "code-with-engineering-playbook.zip",
     "CC BY 4.0"),
    ("Engineering Handbooks", "Google engineering practices", "https://github.com/google/eng-practices/archive/HEAD.zip",
     "eng-practices.zip", "CC BY 3.0"),
    ("Engineering Handbooks", "System design primer", "https://github.com/donnemartin/system-design-primer/archive/HEAD.zip",
     "system-design-primer.zip", "CC BY-SA 4.0"),
    (f"{PUB}/02 Fintech & Payments/Standards & Specs", "UK Open Banking read/write API specs",
     "https://github.com/OpenBankingUK/read-write-api-specs/archive/HEAD.zip", "openbanking-read-write-api-specs.zip", "MIT"),
    (f"{PUB}/02 Fintech & Payments/Standards & Specs", "Berlin Group / FDX style open finance (OpenAPI samples)",
     "https://github.com/APIs-guru/openapi-directory/archive/HEAD.zip", "openapi-directory.zip", "CC0"),
    (f"{PUB}/02 Fintech & Payments/Standards & Specs", "Apache Fineract (core banking) docs and source",
     "https://github.com/apache/fineract/archive/HEAD.zip", "fineract.zip", "Apache-2.0"),
    ("Tool Catalogues", "Awesome game development", "https://github.com/ellisonleao/magictools/archive/HEAD.zip",
     "magictools.zip", "CC BY 4.0"),
    ("Tool Catalogues", "Awesome fintech", "https://github.com/moov-io/awesome-fintech/archive/HEAD.zip",
     "awesome-fintech.zip", "CC0"),
    ("Tool Catalogues", "Awesome self-hosted", "https://github.com/awesome-selfhosted/awesome-selfhosted/archive/HEAD.zip",
     "awesome-selfhosted.zip", "CC BY-SA 3.0"),
    ("Tool Catalogues", "Free for Dev", "https://github.com/ripienaar/free-for-dev/archive/HEAD.zip", "free-for-dev.zip",
     "MIT"),
    ("Tool Catalogues", "Public APIs", "https://github.com/public-apis/public-apis/archive/HEAD.zip", "public-apis.zip",
     "MIT"),
    (f"{PUB}/01 Game Development/Esports Data", "Esports data pipeline (HerrKurz)",
     "https://github.com/HerrKurz/Esports_Data_Pipeline/archive/HEAD.zip", "Esports_Data_Pipeline.zip", "MIT"),
    ("Technology Adoption", "Companies by technology stack (Leadita)",
     "https://github.com/leadita/tech-stack-datasets/archive/HEAD.zip", "tech-stack-datasets.zip",
     "No licence file; README says free to use, provided as-is (domains + detected technologies, no contacts)"),
]


def lifecycles(sink, st):
    """Release and end-of-life dates for ~400 products (OSes, databases, runtimes, devices, cloud services)."""
    stamp = st.get("eol")
    if stamp and stamp > (dt.date.today() - dt.timedelta(days=30)).isoformat():
        return 0
    s = session()
    r = get(s, "https://endoflife.date/api/all.json")
    if not r:
        return 0
    rows = []
    for p in r.json():
        time.sleep(0.3)
        pr = get(s, f"https://endoflife.date/api/{p}.json", tries=3)
        if not pr:
            continue
        for c in pr.json():
            rows.append([p, str(c.get("cycle", "")), c.get("codename"), c.get("releaseDate"), str(c.get("lts", "")),
                         str(c.get("support", "")), str(c.get("eol", "")), str(c.get("extendedSupport", "")),
                         c.get("latest"), c.get("latestReleaseDate"), c.get("link")])
    today = dt.date.today().isoformat()
    wb = new_book()
    add_sheet(wb, "All release cycles", ["Product", "Cycle", "Codename", "Released", "LTS", "Active support until",
                                         "End of life", "Extended support", "Latest version", "Latest release",
                                         "Release notes"], rows, {"Product": 26, "Release notes": 50})
    eol_soon = [x for x in rows if x[6][:4].isdigit() and today <= x[6] <= str(dt.date.today().year + 1) + today[4:]]
    add_sheet(wb, "EOL in next 12 months", ["Product", "Cycle", "Codename", "Released", "LTS", "Active support until",
                                            "End of life", "Extended support", "Latest version", "Latest release",
                                            "Release notes"], eol_soon, {"Product": 26})
    base = f"{TECH}/Technology Lifecycles"
    save_book(sink, f"{base}/Product release and end-of-life dates {today}.xlsx", wb)
    meta_note(sink, base, {"source": "endoflife.date", "url": "https://endoflife.date", "license": "MIT"})
    st["eol"] = today
    return len(rows)


def stackexchange(sink, st):
    """Stack Exchange data dumps (archive.org, CC BY-SA). Files over 3 GB (Stack Overflow) stream straight to Drive."""
    done = set(st.setdefault("se_files", []))
    s = session()
    r = get(s, "https://archive.org/metadata/stackexchange")
    files = {f["name"]: int(f.get("size") or 0) for f in (r.json().get("files", []) if r else [])
             if f.get("name", "").endswith(".7z")}
    n = 0
    for site in SE_SITES:
        for name in sorted(k for k in files if k == f"{site}.7z" or k.startswith(f"{site}-")):
            if name in done or sink.full():
                continue
            url = f"https://archive.org/download/stackexchange/{name}"
            label = SE_LABEL.get(site.split(".")[0], site.split(".")[0].title())
            rel = f"{QA}/{label}/{name}"
            ok = sink.copyurl(url, rel, files[name]) if files[name] > 3 * 2 ** 30 else                 _download(s, url, 3 * 2 ** 30, sink, rel)
            if ok:
                done.add(name)
                n += 1
    st["se_files"] = sorted(done)
    if n:
        meta_note(sink, QA, {"source": "Stack Exchange data dump (Internet Archive)",
                             "url": "https://archive.org/details/stackexchange",
                             "license": "CC BY-SA (Stack Exchange contributors)",
                             "format": "7z of XML: Posts, Comments, Users, Tags, Votes, Badges, PostHistory, PostLinks"})
    return n


SE_LABEL = {"gamedev": "Game Development", "stackoverflow": "Stack Overflow", "softwareengineering": "Software Engineering",
            "codereview": "Code Review", "money": "Personal Finance & Money", "quant": "Quantitative Finance",
            "law": "Law", "workplace": "The Workplace (HR)", "pm": "Project Management", "ux": "User Experience",
            "economics": "Economics", "security": "Information Security", "crypto": "Cryptography",
            "datascience": "Data Science", "dba": "Databases", "devops": "DevOps", "serverfault": "Server Fault",
            "superuser": "Super User", "unix": "Unix & Linux", "webapps": "Web Applications",
            "freelancing": "Freelancing", "startups": "Startups", "patents": "Patents", "opensource": "Open Source",
            "bitcoin": "Bitcoin", "ethereum": "Ethereum", "boardgames": "Board & Card Games", "gaming": "Gaming"}


def docs(sink, st):
    s = session()
    n = 0
    for folder, title, url, name, lic in DOCS:
        if sink.full():
            break
        key = f"doc_{name}"
        stamp = st.get(key)
        if stamp and stamp > (dt.date.today() - dt.timedelta(days=90)).isoformat():
            continue
        dest = f"{folder}/{title}" if folder.startswith(PUB) else f"{TECH}/{folder}/{title}"
        if _download(s, url, 3 * 2 ** 30, sink, f"{dest}/{dt.date.today()} {name}"):
            meta_note(sink, dest, {"source": title, "url": url, "license": lic})
            n += 1
        st[key] = dt.date.today().isoformat()
    return n


# Whole Zenodo records kept together: (record id, destination folder)
RECORDS = [(4595826, f"{PUB}/04 Legal & Compliance/Contracts (CUAD dataset)")]


def records(sink, st):
    """Large curated corpora downloaded file by file (Loghub: 16 system logs, CC BY 4.0)."""
    s = session()
    done = st.setdefault("records_done", [])
    n = 0
    for rid, dest in RECORDS:
        r = get(s, f"https://zenodo.org/api/records/{rid}")
        if not r:
            continue
        d = r.json()
        for f in d.get("files", []):
            key = f"{rid}/{f['key']}"
            if key in done or sink.full():
                continue
            if _download(s, f["links"]["self"], 3 * 2 ** 30, sink, f"{dest}/{f['key']}"):
                n += 1
            done.append(key)
        if not st.get(f"records_meta_{rid}"):
            meta_note(sink, dest, {"source": d["metadata"]["title"], "record": d["links"]["self_html"],
                                   "doi": d.get("doi"), "license": (d["metadata"].get("license") or {}).get("id")})
            st[f"records_meta_{rid}"] = True
    return n


CIC_BUCKET = "https://cse-cic-ids2018.s3.ca-central-1.amazonaws.com"
CIC_PREFIX = "Processed Traffic Data for ML Algorithms/"


def cicids2018(sink, st):
    """CSE-CIC-IDS2018 labelled network flows (CICFlowMeter CSVs, one per capture day; University of New Brunswick)."""
    s = session()
    r = get(s, CIC_BUCKET + "/", params={"list-type": "2", "prefix": CIC_PREFIX})
    if not r:
        return 0
    keys = [k for k in re.findall(r"<Key>(.*?)</Key>", r.text) if k.endswith(".csv")]
    done = st.setdefault("cic2018_done", [])
    n = 0
    for k in keys:
        if k in done or sink.full():
            continue
        name = k.rsplit("/", 1)[1]
        if _download(s, f"{CIC_BUCKET}/{quote(k)}", 6 * 2 ** 30, sink,
                     f"{TI}/Intrusion Detection/CSE-CIC-IDS2018 network flows/{name}"):
            n += 1
            done.append(k)
    if n:
        meta_note(sink, f"{TI}/Intrusion Detection/CSE-CIC-IDS2018 network flows", {
            "source": "CSE-CIC-IDS2018 (Communications Security Establishment & Canadian Institute for Cybersecurity)",
            "url": "https://www.unb.ca/cic/datasets/ids-2018.html", "bucket": CIC_BUCKET,
            "license": "Free for research; cite Sharafaldin, Lashkari and Ghorbani, ICISSP 2018"})
    return n


DOC_PAGES = [  # single web pages kept as DOCX
    ("Developer onboarding plan template (gitmore)",
     "https://gitmore.io/resources/templates/developer-onboarding-plan-template", "gitmore.io free template"),
]


def it_docs(sink, st):
    """IT documentation tools and templates: GitHub topic it-documentation (licensed repos) + template pages."""
    s = session()
    if os.environ.get("GITHUB_TOKEN"):
        s.headers["Authorization"] = f"Bearer {os.environ['GITHUB_TOKEN']}"
    base = f"{TECH}/IT Documentation & Runbooks"
    done = st.setdefault("itdocs_done", {})
    n = 0
    r = get(s, "https://api.github.com/search/repositories",
            params={"q": "topic:it-documentation", "sort": "stars", "per_page": 100})
    for repo in (r.json().get("items", []) if r else []):
        if sink.full():
            break
        lic = (repo.get("license") or {}).get("spdx_id")
        if not lic or lic == "NOASSERTION" or repo.get("size", 0) > 500 * 1024:
            continue
        if done.get(repo["full_name"]) == repo.get("pushed_at"):
            continue
        name = repo["full_name"].replace("/", " - ")
        if _download(s, f"https://github.com/{repo['full_name']}/archive/HEAD.zip", 2 ** 30, sink,
                     f"{base}/IT documentation tools/{name}.zip"):
            meta_note(sink, f"{base}/IT documentation tools/{name}", {
                "source": repo["html_url"], "description": repo.get("description"), "license": lic,
                "stars": repo.get("stargazers_count")})
            n += 1
        done[repo["full_name"]] = repo.get("pushed_at")
    for title, url, credit in DOC_PAGES:
        if done.get(url) or sink.full():
            continue
        pr = get(s, url, headers={"User-Agent": "Mozilla/5.0 (research collector)"})
        if pr:
            doc = _wiki_doc(title, pr.text, url, credit)
            p = sink.path(f"{base}/Templates/{safe(title, 100)}.docx")
            doc.save(p)
            sink.wrote(p)
            n += 1
            done[url] = dt.date.today().isoformat()
    return n


SOURCES = {"records": records, "itdocs": it_docs, "eol": lifecycles, "stackexchange": stackexchange, "techdocs": docs}
