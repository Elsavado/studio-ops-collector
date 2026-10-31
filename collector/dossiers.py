"""Open-source game dossiers: for each open-source game with a public release, everything published about it -
official trailers and screenshots, store and distribution figures (Steam, SteamSpy, Flathub, GitHub releases), news
and patch notes, English player reviews, its Wikipedia article and its art / 3D asset repositories. English only."""
import datetime as dt
import gzip
import csv
import os
import re
import shutil
import subprocess
import time

from .common import add_sheet, get, new_book, safe, save_book, session
from .library import _wiki_doc
from .public import GAME, _download, meta_note

DOS = f"{GAME}/Open-Source Game Dossiers"
SURVEYS = f"{GAME}/Player & Hardware Surveys"

# name: (GitHub repo, Steam store name or None, Flathub id or None, Wikipedia title or None, asset repos)
GAMES = {
    "SuperTuxKart": ("supertuxkart/stk-code", "SuperTuxKart", "net.supertuxkart.SuperTuxKart", "SuperTuxKart",
                     ["supertuxkart/stk-assets-mobile"]),
    "Teeworlds": ("teeworlds/teeworlds", "Teeworlds", "com.teeworlds.Teeworlds", "Teeworlds", []),
    "Hextris": ("Hextris/hextris", None, None, None, []),
    "Godot Tactical RPG": ("ramaureirac/godot-tactical-rpg", None, None, None, []),
    "Mindustry": ("Anuken/Mindustry", "Mindustry", "com.github.Anuken.Mindustry", "Mindustry", ["Anuken/MindustryMaps"]),
    "Shattered Pixel Dungeon": ("00-Evan/shattered-pixel-dungeon", "Shattered Pixel Dungeon",
                                "com.shatteredpixel.shatteredpixeldungeon", "Pixel Dungeon", []),
    "Unciv": ("yairm210/Unciv", "Unciv", "io.github.yairm210.unciv", "Unciv", []),
    "The Battle for Wesnoth": ("wesnoth/wesnoth", "The Battle for Wesnoth", "org.wesnoth.Wesnoth",
                               "The Battle for Wesnoth", ["wesnoth/wesnoth-music"]),
    "OpenTTD": ("OpenTTD/OpenTTD", "OpenTTD", "org.openttd.OpenTTD", "OpenTTD", ["OpenTTD/OpenGFX", "OpenTTD/OpenSFX"]),
    "0 A.D.": ("0ad/0ad", None, "com.play0ad.zeroad", "0 A.D. (video game)", []),
    "Endless Sky": ("endless-sky/endless-sky", "Endless Sky", "io.github.endless_sky.endless_sky", "Endless Sky",
                    ["endless-sky/endless-sky-high-dpi"]),
    "Warzone 2100": ("Warzone2100/warzone2100", "Warzone 2100", "net.wz2100.wz2100", "Warzone 2100", []),
    "Luanti": ("luanti-org/luanti", "Luanti", "org.luanti.luanti", "Luanti", ["luanti-org/minetest_game"]),
    "Red Eclipse": ("redeclipse/base", "Red Eclipse 2", "net.redeclipse.RedEclipse", "Red Eclipse", []),
    "Thrive": ("Revolutionary-Games/Thrive", "Thrive", "com.revolutionarygamesstudio.Thrive", "Thrive (video game)", []),
    "Cataclysm: Dark Days Ahead": ("CleverRaven/Cataclysm-DDA", "Cataclysm: Dark Days Ahead",
                                   "org.cataclysmdda.CataclysmDDA", "Cataclysm: Dark Days Ahead", []),
    "OpenRA": ("OpenRA/OpenRA", None, "net.openra.OpenRA", "OpenRA", []),
    "SuperTux": ("SuperTux/supertux", "SuperTux", "org.supertuxproject.SuperTux", "SuperTux", []),
    "Naev": ("naev/naev", "Naev", "org.naev.Naev", "Naev", ["naev/naev-artwork-production"]),
    "Pioneer": ("pioneerspacesim/pioneer", None, "net.pioneerspacesim.Pioneer", "Pioneer (video game)",
                ["pioneerspacesim/pioneer-assets"]),
    "Zero-K": ("ZeroK-RTS/Zero-K", "Zero-K", None, "Zero-K", []),
    "Beyond All Reason": ("beyond-all-reason/Beyond-All-Reason", None, None, None, []),
    "Hedgewars": ("hedgewars/hw", None, "org.hedgewars.Hedgewars", "Hedgewars", []),
    "Freeciv": ("freeciv/freeciv", None, "org.freeciv.gtk3", "Freeciv", []),
    "Veloren": (None, None, "net.veloren.veloren", "Veloren", []),
    "Xonotic": (None, None, "org.xonotic.Xonotic", "Xonotic", []),
}


def _gh():
    s = session()
    if os.environ.get("GITHUB_TOKEN"):
        s.headers["Authorization"] = f"Bearer {os.environ['GITHUB_TOKEN']}"
    return s


def _steam_id(s, name):
    r = get(s, "https://store.steampowered.com/api/storesearch/", params={"term": name, "cc": "us", "l": "english"})
    for it in ((_json(r) or {}).get("items", [])):
        if it.get("type") == "app" and it["name"].lower().replace(":", "") == name.lower().replace(":", ""):
            return it["id"]
    return None


def _json(r, default=None):
    try:
        return r.json() if r is not None else default
    except ValueError:
        return default


def _cell(v):
    if isinstance(v, dict):
        return ", ".join(f"{k} ({x})" for k, x in v.items())
    if isinstance(v, (list, tuple)):
        return ", ".join(map(str, v))
    return v


def _hls_to_mp4(url, path):
    """Steam trailers are HLS streams; ffmpeg copies the best stream into an MP4 without re-encoding."""
    if not shutil.which("ffmpeg"):
        return False
    r = subprocess.run(["ffmpeg", "-loglevel", "error", "-y", "-i", url, "-c", "copy", "-bsf:a", "aac_adtstoasc", path],
                       timeout=1800)
    return r.returncode == 0 and os.path.exists(path) and os.path.getsize(path) > 10000


def dossiers(sink, st):
    """Per game: trailers, screenshots, store & distribution snapshot, news, reviews, Wikipedia, asset repos."""
    s = session()
    s.headers["User-Agent"] = "Mozilla/5.0 (research collector)"
    gh = _gh()
    state = st.setdefault("dossier", {})
    today = dt.date.today()
    month = today.strftime("%Y-%m")
    n = 0
    for name, (repo, steam_name, flathub, wiki, assets) in GAMES.items():
        if sink.full():
            break
        g = state.setdefault(name, {})
        base = f"{DOS}/{safe(name, 60)}"
        if steam_name and "appid" not in g:
            g["appid"] = _steam_id(s, steam_name)
        appid = g.get("appid")
        # ---- monthly store and distribution snapshot
        if g.get("snap") != month:
            rows, daily, rel = [], [], []
            details = {}
            if appid:
                r = get(s, "https://store.steampowered.com/api/appdetails", params={"appids": appid, "l": "english",
                                                                                     "cc": "us"})
                details = ((_json(r) or {}).get(str(appid)) or {}).get("data", {}) if r else {}
                if details:
                    rows += [["Steam", "App id", appid], ["Steam", "Release date", (details.get("release_date") or {}).get("date")],
                             ["Steam", "Price", (details.get("price_overview") or {}).get("final_formatted", "Free")],
                             ["Steam", "Platforms", ", ".join(k for k, v in (details.get("platforms") or {}).items() if v)],
                             ["Steam", "Genres", ", ".join(x["description"] for x in details.get("genres", []))],
                             ["Steam", "Categories", ", ".join(x["description"] for x in details.get("categories", []))],
                             ["Steam", "Developers", ", ".join(details.get("developers", []))],
                             ["Steam", "Publishers", ", ".join(details.get("publishers", []))],
                             ["Steam", "Recommendations", (details.get("recommendations") or {}).get("total")],
                             ["Steam", "Metacritic", (details.get("metacritic") or {}).get("score")],
                             ["Steam", "Supported languages", re.sub("<[^>]+>", "", details.get("supported_languages", ""))[:500]],
                             ["Steam", "Short description", details.get("short_description")]]
                p = get(s, "https://api.steampowered.com/ISteamUserStats/GetNumberOfCurrentPlayers/v1/",
                        params={"appid": appid})
                if p:
                    rows.append(["Steam", f"Players online at {dt.datetime.utcnow():%Y-%m-%d %H:%M} UTC",
                                 ((_json(p) or {}).get("response") or {}).get("player_count")])
                sp = get(s, "https://steamspy.com/api.php", params={"request": "appdetails", "appid": appid})
                if sp:
                    d = _json(sp) or {}
                    rows += [["SteamSpy", k, d.get(k)] for k in ("owners", "positive", "negative", "average_forever",
                                                                 "median_forever", "ccu", "tags") if k in d]
            if flathub:
                f = get(s, f"https://flathub.org/api/v2/stats/{flathub}")
                if f:
                    d = _json(f) or {}
                    rows.append(["Flathub", "Installs (total)", d.get("installs_total")])
                    daily = sorted((k, v) for k, v in (d.get("installs_per_day") or {}).items())
            if repo:
                r = get(gh, f"https://api.github.com/repos/{repo}")
                if r:
                    d = _json(r) or {}
                    rows += [["GitHub", k, d.get(k)] for k in ("stargazers_count", "forks_count", "open_issues_count",
                                                               "subscribers_count", "created_at", "pushed_at")]
                    rows.append(["GitHub", "Licence", (d.get("license") or {}).get("spdx_id")])
                for page in range(1, 6):
                    r = get(gh, f"https://api.github.com/repos/{repo}/releases", params={"per_page": 100, "page": page})
                    items = _json(r, []) or []
                    for x in items:
                        for a in x.get("assets", []):
                            rel.append([x.get("tag_name"), (x.get("published_at") or "")[:10], a["name"],
                                        a.get("download_count"), round(a.get("size", 0) / 2 ** 20, 1)])
                    if len(items) < 100:
                        break
            rows = [[a, b, _cell(v)] for a, b, v in rows]
            wb = new_book()
            add_sheet(wb, "Snapshot", ["Source", "Measure", "Value"], rows, {"Measure": 40, "Value": 80})
            if daily:
                add_sheet(wb, "Flathub installs per day", ["Date", "Installs"], daily)
            if rel:
                add_sheet(wb, "GitHub release downloads", ["Release", "Published", "Asset", "Downloads", "MB"], rel,
                          {"Asset": 50})
            save_book(sink, f"{base}/Store & Distribution/Distribution snapshot {month}.xlsx", wb, "Distribution",
                      "Public store, download and repository figures")
            g["snap"] = month
            n += 1
            # ---- trailers and screenshots (once each)
            got = set(g.setdefault("media", []))
            for mv in details.get("movies", []) if details else []:
                key = f"m{mv['id']}"
                if key in got or sink.full():
                    continue
                url = mv.get("hls_h264") or (mv.get("mp4") or {}).get("max")
                if not url:
                    continue
                rel_p = f"{base}/Trailers & Videos/{safe(mv.get('name') or str(mv['id']), 80)} ({mv['id']}).mp4"
                p = sink.path(rel_p)
                ok = _hls_to_mp4(url, p) if url.endswith((".m3u8",)) or "m3u8" in url else \
                    _download(s, url, 2 * 2 ** 30, sink, rel_p)
                if ok and "m3u8" in url:
                    sink.wrote(p)
                elif os.path.exists(p) and "m3u8" in url:
                    os.remove(p)
                if ok:
                    got.add(key)
                    n += 1
            for shot in details.get("screenshots", []) if details else []:
                key = f"s{shot['id']}"
                if key in got or sink.full():
                    continue
                url = shot.get("path_full", "").split("?")[0]
                if url and _download(s, url, 50 * 2 ** 20, sink, f"{base}/Screenshots/{shot['id']:03d} "
                                                                    f"{url.rsplit('/', 1)[-1]}"):
                    got.add(key)
            g["media"] = sorted(got)
        # ---- news (Steam announcements + GitHub release notes), refreshed monthly
        if g.get("news") != month and not sink.full():
            news = []
            if appid:
                r = get(s, "https://api.steampowered.com/ISteamNews/GetNewsForApp/v2/",
                        params={"appid": appid, "count": 1000, "maxlength": 0, "format": "json"})
                for x in (((_json(r) or {}).get("appnews") or {}).get("newsitems", [])):
                    news.append([dt.datetime.utcfromtimestamp(x.get("date", 0)).strftime("%Y-%m-%d"), "Steam news",
                                 x.get("feedlabel"), x.get("title"), x.get("author"), x.get("url"),
                                 re.sub(r"\[/?[^\]]+\]", "", x.get("contents") or "")[:32000]])
            if repo:
                for page in range(1, 4):
                    r = get(gh, f"https://api.github.com/repos/{repo}/releases", params={"per_page": 100, "page": page})
                    items = _json(r, []) or []
                    for x in items:
                        news.append([(x.get("published_at") or "")[:10], "Release notes", x.get("tag_name"),
                                     x.get("name") or x.get("tag_name"), (x.get("author") or {}).get("login"),
                                     x.get("html_url"), (x.get("body") or "")[:32000]])
                    if len(items) < 100:
                        break
            if news:
                news.sort(reverse=True)
                wb = new_book()
                add_sheet(wb, "News & patch notes", ["Date", "Type", "Feed / tag", "Title", "Author", "URL", "Text"],
                          news, {"Title": 50, "URL": 40, "Text": 100})
                save_book(sink, f"{base}/News & Patch Notes/News and patch notes (to {today}).xlsx", wb, "News",
                          "Steam announcements and GitHub release notes")
                n += 1
            g["news"] = month
        # ---- English player reviews, refreshed every three months
        if appid and (not g.get("reviews") or g["reviews"] < (today - dt.timedelta(days=90)).isoformat()) \
                and not sink.full():
            rows, cursor, summary = [], "*", {}
            for _ in range(200):  # up to 20,000 reviews
                r = get(s, f"https://store.steampowered.com/appreviews/{appid}",
                        params={"json": 1, "language": "english", "filter": "recent", "purchase_type": "all",
                                "num_per_page": 100, "cursor": cursor})
                d = _json(r, {}) or {}
                if not summary:
                    summary = d.get("query_summary") or {}
                revs = d.get("reviews") or []
                for x in revs:
                    a = x.get("author") or {}
                    rows.append([x.get("recommendationid"), dt.datetime.utcfromtimestamp(x.get("timestamp_created", 0)).date(),
                                 "Recommended" if x.get("voted_up") else "Not recommended",
                                 round((a.get("playtime_forever") or 0) / 60, 1),
                                 round((a.get("playtime_at_review") or 0) / 60, 1), a.get("num_games_owned"),
                                 x.get("votes_up"), x.get("votes_funny"), x.get("steam_purchase"),
                                 x.get("received_for_free"), x.get("written_during_early_access"),
                                 (x.get("review") or "").replace("\r", " ")])
                cursor = d.get("cursor")
                if not revs or not cursor:
                    break
                time.sleep(1.2)
            if rows:
                p = sink.path(f"{base}/Player Reviews/Steam reviews (English, to {today}).csv.gz")
                with gzip.open(p, "wt", encoding="utf-8", newline="") as f:
                    w = csv.writer(f)
                    w.writerow(["Review", "Date", "Verdict", "Playtime hours", "Playtime at review hours",
                                "Games owned", "Helpful votes", "Funny votes", "Bought on Steam", "Got for free",
                                "Early access", "Text"])
                    w.writerows(rows)
                sink.wrote(p)
                wb = new_book()
                add_sheet(wb, "Review summary", ["Measure", "Value"], [[k, _cell(v)] for k, v in summary.items()],
                          {"Measure": 30, "Value": 30})
                save_book(sink, f"{base}/Player Reviews/Review summary {today}.xlsx", wb, "Reviews", "Steam")
                n += 1
                g["reviews"] = today.isoformat()  # empty answers are retried next run
        # ---- Wikipedia article (once a year)
        if wiki and (not g.get("wiki") or g["wiki"] < (today - dt.timedelta(days=365)).isoformat()) and not sink.full():
            r = get(s, f"https://en.wikipedia.org/api/rest_v1/page/html/{wiki.replace(' ', '_')}")
            if r:
                doc = _wiki_doc(wiki, r.text, f"https://en.wikipedia.org/wiki/{wiki.replace(' ', '_')}")
                p = sink.path(f"{base}/About/{safe(wiki, 80)} (Wikipedia).docx")
                doc.save(p)
                sink.wrote(p)
                g["wiki"] = today.isoformat()
        # ---- art, music and 3D asset repositories (archive, refreshed yearly)
        for a in assets:
            key = f"asset_{a}"
            if sink.full() or g.get(key, "") > (today - dt.timedelta(days=365)).isoformat():
                continue
            r = get(gh, f"https://api.github.com/repos/{a}")
            info = _json(r, {}) or {}
            if not info or info.get("size", 0) > 6 * 2 ** 20:  # KB; skip archives over ~6 GB
                g[key] = today.isoformat()
                continue
            url = f"https://codeload.github.com/{a}/zip/refs/heads/{info.get('default_branch', 'master')}"
            rel_p = f"{base}/Art, 3D & Audio Assets/{safe(a.replace('/', ' - '), 80)}.zip"
            if info.get("size", 0) > 2.5 * 2 ** 20:
                ok = sink.copyurl(url, rel_p, info["size"] * 1024)
            else:
                ok = _download(s, url, 3 * 2 ** 30, sink, rel_p)
            if ok:
                n += 1
                meta_note(sink, f"{base}/Art, 3D & Audio Assets", {"source": info.get("html_url"),
                                                                    "license": (info.get("license") or {}).get("spdx_id")})
            g[key] = today.isoformat()
    meta_note(sink, DOS, {"source": "Steam store/news/reviews APIs, SteamSpy, Flathub, GitHub, Wikipedia",
                          "note": "Per-game public data; English only"})
    return n


def hw_survey(sink, st):
    """Steam Hardware & Software Survey, one PDF per month (what hardware, OS and languages players use)."""
    month = dt.date.today().strftime("%Y-%m")
    if st.get("hw_survey") == month:
        return 0
    s = session()
    s.headers["User-Agent"] = "Mozilla/5.0 (research collector)"
    n = 0
    for part, url in [("Combined", "https://store.steampowered.com/hwsurvey/?platform=combined"),
                      ("Windows", "https://store.steampowered.com/hwsurvey/?platform=pc"),
                      ("Linux", "https://store.steampowered.com/hwsurvey/?platform=linux"),
                      ("Mac", "https://store.steampowered.com/hwsurvey/?platform=mac"),
                      ("Video cards (DirectX)", "https://store.steampowered.com/hwsurvey/videocard/"),
                      ("Processors", "https://store.steampowered.com/hwsurvey/processormfg/")]:
        if _download(s, url, 50 * 2 ** 20, sink, f"{SURVEYS}/Steam Hardware & Software Survey/{month}/{part}.html"):
            n += 1
    st["hw_survey"] = month
    return n


SOURCES = {"dossiers": dossiers, "hwsurvey": hw_survey}
