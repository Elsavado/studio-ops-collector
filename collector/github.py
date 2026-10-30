"""GitHub: technical repositories, Jupyter notebooks and open-source apps for a security & IT services business.

Only licensed repositories are taken. Everything is filed by topic, never by where it came from."""
import os
import re
import time

from .common import get, safe, session
from .public import PUB, _download

REPOS = f"{PUB}/11 GitHub Repositories"
NOTEBOOKS = f"{PUB}/12 Jupyter Notebooks"
APPS = f"{PUB}/13 Software & Apps"

REPO_TOPICS = {  # folder: GitHub topics
    "Game Engines & Frameworks": ["game-engine", "godot", "unity3d", "phaser", "bevy", "monogame", "libgdx", "raylib"],
    "Game Development": ["gamedev", "game-development", "indie-game", "game-jam", "procedural-generation",
                         "game-design", "level-editor"],
    "Multiplayer & Game Backends": ["multiplayer", "game-server", "netcode", "matchmaking", "mmo", "websocket-game"],
    "Game Analytics & Live Ops": ["game-analytics", "player-analytics", "liveops", "telemetry"],
    "Payments & Fintech": ["fintech", "payments", "payment-gateway", "stripe", "open-banking", "banking", "neobank",
                           "psd2"],
    "Ledgers & Accounting": ["ledger", "accounting", "double-entry-accounting", "bookkeeping", "invoicing",
                             "billing"],
    "Trading & Crypto": ["trading", "algorithmic-trading", "cryptocurrency", "blockchain", "defi", "exchange"],
    "Fraud, Risk & Compliance": ["fraud-detection", "kyc", "aml", "risk-management", "compliance"],
    "HR & People Ops": ["hrms", "human-resources", "payroll", "recruitment", "applicant-tracking-system"],
    "Legal Tech": ["legal-tech", "contract-management", "legaltech", "e-signature"],
    "Project Management & Collaboration": ["project-management", "kanban", "issue-tracker", "wiki", "documentation"],
    "CRM, Marketing & Support": ["crm", "marketing-automation", "helpdesk", "customer-support", "email-marketing"],
    "DevOps & Cloud": ["devops", "ci-cd", "kubernetes", "infrastructure-as-code", "observability", "self-hosted"],
    "Web & Mobile Development": ["react-native", "flutter", "nextjs", "android-app", "ios-app"],
    "Security & Privacy": ["application-security", "secrets-management", "authentication", "oauth2", "encryption"],
}

NOTEBOOK_TOPICS = {
    "Game & Player Analytics": ["player churn", "game analytics", "video game sales", "steam", "player retention",
                                "game recommendation"],
    "Fraud & Risk": ["fraud detection", "credit card fraud", "anti money laundering", "credit risk", "loan default"],
    "Payments & Fintech Analytics": ["payment", "transaction analysis", "fintech", "bank customer"],
    "Trading & Quant": ["algorithmic trading", "portfolio optimization", "stock prediction", "quantitative finance"],
    "Accounting & Financial Analysis": ["financial statement analysis", "financial analysis", "accounting",
                                        "forecasting revenue"],
    "HR & People Analytics": ["employee attrition", "hr analytics", "salary prediction", "resume screening"],
    "Legal & Contract NLP": ["contract analysis", "legal nlp", "legal text"],
    "Customer & Marketing Analytics": ["customer churn", "customer segmentation", "marketing analytics",
                                       "sentiment analysis reviews"],
    "Product & Experimentation": ["a/b testing", "product analytics", "recommendation system"],
    "Software Engineering Analytics": ["github analysis", "bug prediction", "code analysis"],
}

APPS_LIST = {  # folder: GitHub repositories of open-source apps
    "Game Engines & Tools": ["godotengine/godot", "defold/defold", "mapeditor/tiled", "Orama-Interactive/Pixelorama",
                             "LibreSprite/LibreSprite", "aseprite/aseprite", "blender/blender", "PixiEditor/PixiEditor",
                             "ldtk/ldtk", "armory3d/armorpaint"],
    "Game Backends & Multiplayer": ["heroiclabs/nakama", "colyseus/colyseus", "googleforgames/agones",
                                    "googleforgames/open-match"],
    "Payments & Billing": ["juspay/hyperswitch", "killbill/killbill", "getlago/lago", "formancehq/ledger",
                           "blnkfinance/blnk", "apache/fineract"],
    "Accounting & ERP": ["firefly-iii/firefly-iii", "invoiceninja/invoiceninja", "akaunting/akaunting",
                         "frappe/erpnext", "bigcapitalhq/bigcapital", "beancount/fava", "kresusapp/kresus"],
    "HR & Payroll": ["frappe/hrms", "orangehrm/orangehrm", "opencats/OpenCATS", "horilla-opensource/horilla"],
    "Legal & Documents": ["docusealco/docuseal", "documenso/documenso", "jhpyle/docassemble", "paperless-ngx/paperless-ngx"],
    "Project Management": ["makeplane/plane", "taigaio/taiga-back", "opf/openproject", "kanboard/kanboard",
                           "wekan/wekan", "focalboard/focalboard"],
    "CRM, Support & Marketing": ["twentyhq/twenty", "chatwoot/chatwoot", "SuiteCRM/SuiteCRM", "mautic/mautic",
                                 "zammad/zammad", "listmonk/listmonk"],
    "Knowledge & Collaboration": ["outline/outline", "BookStackApp/BookStack", "wikijs/wiki", "mattermost/mattermost",
                                  "RocketChat/Rocket.Chat", "nextcloud/server"],
    "Analytics & BI": ["metabase/metabase", "apache/superset", "plausible/analytics", "PostHog/posthog",
                       "matomo-org/matomo", "growthbook/growthbook"],
    "Developer Tools & DevOps": ["go-gitea/gitea", "woodpecker-ci/woodpecker", "sentry/self-hosted",
                                 "grafana/grafana", "prometheus/prometheus", "louislam/uptime-kuma", "portainer/portainer",
                                 "hashicorp/terraform-provider-aws", "opentofu/opentofu"],
    "Security & Identity": ["keycloak/keycloak", "goauthentik/authentik", "dani-garcia/vaultwarden",
                            "keepassxreboot/keepassxc", "zitadel/zitadel"],
}
ASSET_OK = re.compile(r"(windows|win64|win-x64|x64\.(msi|exe)|amd64|x86_64|linux64|\.msi$|setup.*\.exe$|\.deb$|"
                      r"\.rpm$|\.appimage$)", re.I)
ASSET_BAD = re.compile(r"(arm|aarch64|i386|i686|386|mips|ppc|s390|riscv|darwin|macos|\.sig$|\.asc$|\.sha\d*|"
                       r"checksums|sbom|\.pem$|\.minisig$|debug|symbols)", re.I)


def _gh():
    s = session()
    if os.environ.get("GITHUB_TOKEN"):
        s.headers["Authorization"] = f"Bearer {os.environ['GITHUB_TOKEN']}"
    s.headers["Accept"] = "application/vnd.github+json"
    return s


def _search(s, q, page):
    time.sleep(2.2)  # search API: 30 requests a minute
    r = get(s, "https://api.github.com/search/repositories",
            params={"q": q, "sort": "stars", "order": "desc", "per_page": 30, "page": page})
    return r.json().get("items", []) if r else []


def _licensed(repo):
    lic = (repo.get("license") or {}).get("spdx_id")
    return lic and lic != "NOASSERTION"


def repos(sink, st, per_topic=3, max_kb=400 * 1024):
    """Default-branch archives of the most-starred licensed repositories per topic."""
    s = _gh()
    done = st.setdefault("gh_repos", {})
    pages = st.setdefault("gh_repo_page", {})
    n = 0
    for folder, topics in REPO_TOPICS.items():
        for t in topics:
            if sink.full():
                return n
            page = pages.get(t, 1)
            items = _search(s, f"topic:{t} archived:false", page)
            pages[t] = page + 1 if items else 1
            took = 0
            for repo in items:
                if took >= per_topic or sink.full():
                    break
                name = repo["full_name"]
                if name in done or not _licensed(repo) or repo.get("size", 0) > max_kb or repo.get("fork"):
                    continue
                dest = f"{REPOS}/{folder}/{safe(name.replace('/', ' - '), 100)}.zip"
                if _download(s, f"https://api.github.com/repos/{name}/zipball", 2 * 2 ** 30, sink, dest):
                    sink.state.data.setdefault("catalog", {})[dest] = {
                        "source": repo["html_url"], "license": repo["license"]["spdx_id"],
                        "stars": repo.get("stargazers_count"), "description": (repo.get("description") or "")[:500]}
                    took += 1
                    n += 1
                done[name] = repo.get("pushed_at")
    return n


def notebooks(sink, st, per_query=3, per_repo=60, max_mb=25):
    """Jupyter notebooks (.ipynb) from licensed notebook repositories, with their folder structure."""
    s = _gh()
    done = st.setdefault("nb_repos", [])
    done_set = set(done)
    pages = st.setdefault("nb_page", {})
    n = 0
    for folder, queries in NOTEBOOK_TOPICS.items():
        for q in queries:
            if sink.full():
                return n
            page = pages.get(q, 1)
            items = _search(s, f'{q} language:jupyter-notebook archived:false', page)
            pages[q] = page + 1 if items else 1
            took = 0
            for repo in items:
                if took >= per_query or sink.full():
                    break
                name = repo["full_name"]
                about = " ".join([name, repo.get("description") or "", " ".join(repo.get("topics") or [])]).lower()
                if name in done_set or not _licensed(repo) or not all(w in about for w in q.lower().split()):
                    continue
                done.append(name)
                done_set.add(name)
                tree = get(s, f"https://api.github.com/repos/{name}/git/trees/{repo['default_branch']}",
                           params={"recursive": "1"})
                files = [f for f in (tree.json().get("tree", []) if tree else [])
                         if f.get("type") == "blob" and f["path"].lower().endswith(".ipynb")
                         and f.get("size", 0) <= max_mb * 2 ** 20]
                got = 0
                for f in files[:per_repo]:
                    url = f"https://raw.githubusercontent.com/{name}/{repo['default_branch']}/{f['path']}"
                    if _download(s, url, max_mb * 2 ** 20, sink,
                                 f"{NOTEBOOKS}/{folder}/{safe(name.replace('/', ' - '), 90)}/{f['path']}"):
                        got += 1
                if got:
                    sink.state.data.setdefault("catalog", {})[f"{NOTEBOOKS}/{folder}/{name}"] = {
                        "source": repo["html_url"], "license": repo["license"]["spdx_id"], "notebooks": got}
                    took += 1
                    n += got
    return n


def apps(sink, st, app_cap=2 * 2 ** 30):
    """Latest release of each app: Windows / Linux x64 installers and packages, else the release source archive.
    A new release replaces nothing; it is added under its own version folder."""
    s = _gh()
    done = st.setdefault("app_release", {})
    n = 0
    for folder, names in APPS_LIST.items():
        for name in names:
            if sink.full():
                return n
            r = get(s, f"https://api.github.com/repos/{name}/releases/latest")
            rel = r.json() if r else None
            if not rel or not rel.get("tag_name"):
                tags = get(s, f"https://api.github.com/repos/{name}/tags", params={"per_page": 1})
                tag = (tags.json() or [{}])[0].get("name") if tags else None
                if not tag:
                    continue
                rel = {"tag_name": tag, "assets": [], "zipball_url": f"https://api.github.com/repos/{name}/zipball/{tag}"}
            tag = rel["tag_name"]
            if done.get(name) == tag:
                continue
            app = name.split("/")[1]
            dest = f"{APPS}/{folder}/{safe(app, 60)}/{safe(tag, 40)}"
            assets = [a for a in rel.get("assets", []) if ASSET_OK.search(a["name"]) and not ASSET_BAD.search(a["name"])]
            total, got = 0, 0
            for a in sorted(assets, key=lambda a: a["size"]):
                total += a["size"]
                if total > app_cap or sink.full():
                    break
                if _download(s, a["browser_download_url"], app_cap, sink, f"{dest}/{a['name']}"):
                    got += 1
            if _download(s, rel.get("zipball_url") or f"https://api.github.com/repos/{name}/zipball/{tag}", app_cap,
                         sink, f"{dest}/{safe(app, 60)}-{safe(tag, 40)}-source.zip"):
                got += 1
            if got:
                done[name] = tag
                sink.state.data.setdefault("catalog", {})[dest] = {"source": f"https://github.com/{name}",
                                                                   "release": tag, "files": got}
                n += 1
    return n


SOURCES = {"ghrepos": repos, "notebooks": notebooks, "apps": apps}
