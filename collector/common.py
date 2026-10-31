"""Shared helpers: HTTP session, run state, staging sink that ships files to Drive with rclone."""
import json
import os
import re
import shutil
import subprocess
import time

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

UA = "ops-data-collector/1.0 (research; github.com/Elsavado/ops-data-collector)"


def session():
    s = requests.Session()
    s.headers["User-Agent"] = UA
    return s


def get(s, url, tries=5, **kw):
    """GET with backoff on 429/5xx and network errors. Returns the response or None."""
    kw.setdefault("timeout", 120)
    for i in range(tries):
        try:
            r = s.get(url, **kw)
            if r.status_code == 200:
                return r
            if r.status_code in (403, 404, 410):
                return None
        except requests.RequestException:
            pass
        time.sleep(min(10 * 2 ** i, 120))
    return None


def safe(name, limit=120):
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', " ", str(name))
    name = re.sub(r"\s+", " ", name).strip(" .")
    return (name[:limit].rstrip(" .") or "untitled")


EN_WORDS = set("the of and to in a is for on with by from this that are as be at or an data dataset using it we "
               "which our can has have was were not all these its new between based".split())
EN_LANG = {"en", "eng", "english", "en-us", "en-gb", "en_us", "en_gb"}


def english(text, lang=None):
    """True when content is English: a stated language must be English; otherwise the text (title + description)
    must be mostly Latin letters with common English words among them."""
    if lang:
        langs = lang if isinstance(lang, (list, tuple, set)) else [lang]
        return any(str(x).strip().lower() in EN_LANG for x in langs)
    text = text or ""
    letters = [c for c in text if c.isalpha()]
    if letters and sum(c.isascii() for c in letters) / len(letters) < 0.9:
        return False
    words = re.findall(r"[a-z']+", text.lower())
    if len(words) < 4:
        return True  # too short to judge (a bare title); ASCII check above already passed
    return sum(w in EN_WORDS for w in words) / len(words) >= 0.08 or len(words) < 8 and any(w in EN_WORDS for w in words)


LANG_CODE = (r"(ar|bn|bg|ca|cs|da|de|el|es|et|fa|fi|fr|he|hi|hu|id|it|ja|ko|lt|lv|ms|nl|no|nb|pl|pt|pt-br|pt_br|ro|"
             r"ru|sk|sl|sr|sv|th|tr|uk|ur|vi|zh|zh-cn|zh-tw|zh_cn|zh_tw|zh-hans|zh-hant|fil|sw|ta|te)")
LANG_DIR = re.compile(LANG_CODE + r"$", re.I)
LANG_PARENTS = {"content", "docs", "doc", "i18n", "locales", "locale", "translations", "translation", "l10n", "lang",
                "languages", "copy", "manuals", "po", "localization", "localisation"}
LANG_FILE = re.compile(r"[-_.]" + LANG_CODE + r"\.(md|mdx|rst|txt|html?|json|ya?ml|po|adoc|tex)$", re.I)


def translated(member):
    """True for an archive member that is a translation: a language folder under content/docs/i18n/..., or a
    file named like README-ja.md / index.zh-CN.md."""
    segs = member.split("/")
    for i in range(1, len(segs) - 1):
        if LANG_DIR.match(segs[i]) and segs[i - 1].lower() in LANG_PARENTS:
            return True
    return bool(LANG_FILE.search(segs[-1]))


def english_zip(path):
    """Rewrite a zip without its translations (keeps English). Returns bytes removed."""
    import zipfile
    with zipfile.ZipFile(path) as zf:
        drop = [i for i in zf.infolist() if translated(i.filename)]
        if not drop:
            return 0
        tmp = path + ".en"
        skip = {i.filename for i in drop}
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for i in zf.infolist():
                if i.filename not in skip:
                    out.writestr(i, zf.read(i.filename), zipfile.ZIP_DEFLATED)
    os.replace(tmp, path)
    return sum(i.file_size for i in drop)


class State:
    def __init__(self, path):
        self.path = path
        self.data = {}
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                self.data = json.load(f)

    def section(self, key):
        return self.data.setdefault(key, {})

    def save(self):
        os.makedirs(os.path.dirname(self.path) or ".", exist_ok=True)
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=1, sort_keys=True)
        os.replace(tmp, self.path)


class Sink:
    """Files are written under out/<drive path>. flush() moves them to the Drive root with rclone
    (live) or just counts and deletes them (dry run). Logs only show counts, never names."""

    def __init__(self, out, remote, live, state, budget_bytes, deadline, keep=False):
        self.out, self.remote, self.live, self.state, self.keep = out, remote, live, state, keep
        self.budget = budget_bytes
        self.deadline = deadline
        self.slice_end = deadline  # per-source time slice, set by the runner
        self.sent_bytes = 0
        self.files = 0
        self.pending = 0
        os.makedirs(out, exist_ok=True)

    def path(self, rel):
        p = os.path.join(self.out, *[safe(x, 150) for x in rel.split("/")])
        os.makedirs(os.path.dirname(p), exist_ok=True)
        return p

    def wrote(self, p):
        n = os.path.getsize(p)
        self.pending += n
        self.files += 1
        if self.pending > 400 * 2 ** 20:
            self.flush()

    def write_bytes(self, rel, data):
        p = self.path(rel)
        with open(p, "wb") as f:
            f.write(data)
        self.wrote(p)

    def write_text(self, rel, text):
        self.write_bytes(rel, text.encode("utf-8"))

    def copyurl(self, url, rel, size=0):
        """Very large files go straight from the source to Drive (never touch the runner's disk)."""
        if self.live:
            dest = self.remote + "/" + "/".join(safe(x, 150) for x in rel.split("/"))
            r = subprocess.run(["rclone", "copyurl", url, dest, "--retries", "3", "--low-level-retries", "20", "-q"])
            if r.returncode != 0:
                return False
        self.sent_bytes += size
        self.files += 1
        return True

    def full(self):
        return self.sent_bytes + self.pending >= self.budget or time.time() > min(self.deadline, self.slice_end)

    def flush(self):
        if self.keep and not self.live:
            self.state.save()
            self.sent_bytes += self.pending
            self.pending = 0
            return
        if self.pending == 0 and not os.listdir(self.out):
            self.state.save()
            return
        if self.live:
            cmd = ["rclone", "move", self.out, self.remote, "--transfers", "4", "--checkers", "8",
                   "--drive-chunk-size", "64M", "--retries", "5", "--low-level-retries", "20",
                   "--delete-empty-src-dirs", "-q"]
            subprocess.run(cmd, check=True)
            self.state.save()
            subprocess.run(["rclone", "copyto", self.state.path,
                            self.remote + "/_state/" + os.path.basename(self.state.path), "-q"],
                           check=True)
        else:
            self.state.save()
            shutil.rmtree(self.out, ignore_errors=True)
        os.makedirs(self.out, exist_ok=True)
        self.sent_bytes += self.pending
        self.pending = 0


HEAD_FILL = PatternFill("solid", fgColor="1F3A5F")
HEAD_FONT = Font(bold=True, color="FFFFFF")


def add_sheet(wb, title, header, rows, widths=None):
    ws = wb.create_sheet(safe(title, 31)[:31])
    ws.append(list(header))
    for c in ws[1]:
        c.fill, c.font = HEAD_FILL, HEAD_FONT
    for r in rows:
        ws.append([("" if v is None else v) for v in r])
    ws.freeze_panes = "A2"
    for i, h in enumerate(header, 1):
        w = (widths or {}).get(h) or min(max(len(str(h)) + 2, 12), 50)
        ws.column_dimensions[get_column_letter(i)].width = w
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    return ws


def new_book():
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def save_book(sink, rel, wb, title=None, note=None):
    if title or note:
        wb.properties.title = title
        wb.properties.description = note
    p = sink.path(rel)
    wb.save(p)
    sink.wrote(p)
