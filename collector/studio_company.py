"""Fictional game & fintech studio "Corvalis Labs" - SYNTHETIC records for research.

Corvalis Labs builds games and fintech software for clients (fixed-price and time-and-materials projects), runs its
own live games and a payments platform (CorvalisPay). The world (clients, projects, products, staff) is simulated month
by month from a fixed seed, so reruns reproduce past months exactly and only add new ones.
"""
import csv
import datetime as dt
import gzip
import io
import math
import random
import sys
import zipfile
from email.message import EmailMessage
from email.utils import format_datetime

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from . import names as N
from . import staff_records as S
from .common import add_sheet, new_book, safe, save_book

CO = "Corvalis Labs"
DOMAIN = "corvalis.example"
ROOT = f"Fictional Company - {CO}"
F = {k: f"{ROOT}/{v}" for k, v in {
    "crm": "01 Clients & CRM", "legal": "02 Contracts & Legal", "proj": "03 Projects & Delivery",
    "eng": "04 Engineering", "games": "05 Live Games & Player Data", "pay": "06 Fintech Platform (CorvalisPay)",
    "support": "07 Customer Support", "fin": "08 Finance & Accounting", "hr": "09 HR & People",
    "mkt": "10 Sales & Marketing", "rep": "11 Management Reports", "pol": "12 Policies & SOPs",
    "mail": "13 Emails & Meetings"}.items()}
BANNER = "FICTIONAL - synthetic research data. Corvalis Labs, its clients, players, merchants and staff are made up."
START = (2018, 1)
SEED = 20180101
SOP_SCOPE = "Engineering, DevOps, Security, Payments operations, Support"
WORK_LABELS = ("Commits", "Issues resolved", "Support tickets")

ACCOUNTS = {
    "1000": ("Cash at bank", "Asset"), "1050": ("Settlement funds held for merchants", "Asset"),
    "1100": ("Accounts receivable", "Asset"), "1150": ("App store receivables", "Asset"),
    "1500": ("Equipment & devices", "Asset"), "1510": ("Accumulated depreciation", "Asset"),
    "2000": ("Accounts payable", "Liability"), "2050": ("Merchant settlement payable", "Liability"),
    "2100": ("Sales tax payable", "Liability"), "2200": ("Accrued payroll taxes & benefits", "Liability"),
    "3000": ("Share capital", "Equity"), "3100": ("Retained earnings", "Equity"),
    "4000": ("Client services revenue", "Revenue"), "4100": ("Game revenue (net of store fees)", "Revenue"),
    "4200": ("Advertising revenue", "Revenue"), "4300": ("Payment processing revenue", "Revenue"),
    "5000": ("Salaries & wages", "Expense"), "5050": ("Payroll taxes & benefits", "Expense"),
    "5100": ("Contractors", "Expense"), "5150": ("Card network & bank fees", "Expense"),
    "5200": ("Cloud hosting", "Expense"), "5300": ("User acquisition & marketing", "Expense"),
    "5400": ("Rent & office", "Expense"), "5500": ("Software & tools", "Expense"),
    "5600": ("Legal, audit & compliance", "Expense"), "5700": ("Insurance", "Expense"),
    "5800": ("Depreciation", "Expense"), "5900": ("Fraud & chargeback losses", "Expense"),
    "6000": ("Training & recruiting", "Expense"), "6200": ("Staff travel & expense claims", "Expense")}

ROLES = {  # role: (department, monthly base USD, required headcount from (projects, players_k, merchants))
    "Chief Executive Officer": ("Executive", 21000, lambda p, g, m: 1),
    "Chief Technology Officer": ("Engineering", 19500, lambda p, g, m: 1),
    "Chief Financial Officer": ("Finance", 18000, lambda p, g, m: 1),
    "Head of Games": ("Games", 15500, lambda p, g, m: 1),
    "Head of Fintech": ("Fintech", 15500, lambda p, g, m: 1),
    "Head of People": ("People", 12500, lambda p, g, m: 1),
    "Engineering Manager": ("Engineering", 14500, lambda p, g, m: 1 + p // 6),
    "Backend Engineer": ("Engineering", 11500, lambda p, g, m: 3 + p + m // 400),
    "Frontend Engineer": ("Engineering", 10500, lambda p, g, m: 2 + p // 2),
    "Mobile Engineer": ("Engineering", 11000, lambda p, g, m: 2 + p // 2),
    "Game Developer": ("Games", 10000, lambda p, g, m: 3 + p // 2 + g // 300),
    "Game Designer": ("Games", 8800, lambda p, g, m: 1 + p // 4 + g // 600),
    "Technical Artist": ("Games", 8500, lambda p, g, m: 1 + p // 4),
    "QA Engineer": ("Engineering", 7600, lambda p, g, m: 2 + p // 3),
    "DevOps Engineer": ("Engineering", 12000, lambda p, g, m: 1 + p // 5 + m // 600),
    "Security Engineer": ("Engineering", 13000, lambda p, g, m: 1 + m // 500),
    "Data Analyst": ("Engineering", 9000, lambda p, g, m: 1 + g // 400 + m // 700),
    "Product Manager": ("Product", 11500, lambda p, g, m: 1 + p // 4),
    "Project Manager": ("Delivery", 9500, lambda p, g, m: 1 + p // 3),
    "UX Designer": ("Product", 9000, lambda p, g, m: 1 + p // 4),
    "Compliance Officer": ("Fintech", 10500, lambda p, g, m: 1 if m else 0),
    "Payments Operations Analyst": ("Fintech", 7200, lambda p, g, m: (1 + m // 250) if m else 0),
    "Customer Support Agent": ("Support", 4800, lambda p, g, m: 2 + g // 150 + m // 200),
    "Account Executive": ("Sales", 9000, lambda p, g, m: 1 + p // 4),
    "Marketing Manager": ("Marketing", 9200, lambda p, g, m: 1 + g // 500),
    "Finance Officer": ("Finance", 7000, lambda p, g, m: 1 + p // 8),
    "Recruiter": ("People", 6800, lambda p, g, m: 1 + p // 10),
    "Legal Counsel": ("Legal", 14000, lambda p, g, m: 1 if p > 4 or m else 0),
}
HEADS = {"Executive": "Chief Executive Officer", "Engineering": "Chief Technology Officer",
         "Finance": "Chief Financial Officer", "Games": "Head of Games", "Fintech": "Head of Fintech",
         "People": "Head of People", "Product": "Chief Technology Officer", "Delivery": "Chief Technology Officer",
         "Support": "Head of Games", "Sales": "Chief Executive Officer", "Marketing": "Chief Executive Officer",
         "Legal": "Chief Executive Officer"}
EXPENSES = {
    "Account Executive": [("Client entertainment", "Dinner with client", 80, 600, .5), ("Travel", "Client pitch trip", 300, 2600, .3)],
    "Project Manager": [("Travel", "Client workshop", 250, 1800, .25), ("Taxi", "Client meeting", 15, 90, .4)],
    "Game Designer": [("Conferences", "Game conference ticket", 300, 1600, .08), ("Software", "Design tools", 15, 80, .2)],
    "Game Developer": [("Hardware", "Test device", 150, 900, .08), ("Software", "Asset store purchase", 10, 120, .2)],
    "Mobile Engineer": [("Hardware", "Test phone", 200, 1100, .07)],
    "Backend Engineer": [("Conferences", "Conference ticket", 300, 1500, .05), ("Books & courses", "Online course", 20, 200, .15)],
    "Security Engineer": [("Certification", "Security certification exam", 300, 900, .07)],
    "Marketing Manager": [("Events", "Booth materials", 200, 2500, .2), ("Travel", "Trade show", 400, 2400, .2)],
    "Recruiter": [("Events", "Career fair", 150, 900, .2)],
}
TRAINING = {
    "Backend Engineer": ["Secure coding (OWASP Top 10)", "Distributed systems", "Kubernetes for developers"],
    "Frontend Engineer": ["Accessibility (WCAG)", "Web performance"], "Mobile Engineer": ["Mobile app security (MASVS)"],
    "Game Developer": ["Multiplayer netcode", "Engine profiling", "Shader programming"],
    "Game Designer": ["Free-to-play economy design", "Level design workshop"],
    "QA Engineer": ["Test automation", "Exploratory testing"], "DevOps Engineer": ["Infrastructure as code", "Incident command"],
    "Security Engineer": ["PCI DSS for engineers", "Threat modelling"], "Compliance Officer": ["AML/KYC certification", "PCI DSS"],
    "Payments Operations Analyst": ["Chargeback management", "Card network rules"],
    "Customer Support Agent": ["De-escalation", "Payments support basics"], "Product Manager": ["Product discovery"],
}
COMMON_TRAINING = ["Security awareness", "Privacy & data protection", "Anti-harassment"]

GAMES = [  # name, genre, platform, launch, peak DAU (thousands), ARPDAU USD
    ("Skyforge Tactics", "Strategy RPG", "Mobile", (2019, 3), 240, 0.21),
    ("Neon Drift", "Racing", "PC", (2020, 9), 85, 0.34),
    ("Puzzle Harbor", "Casual puzzle", "Web & Mobile", (2021, 5), 410, 0.07),
    ("Ironvale Arena", "Multiplayer battler", "PC & Console", (2023, 10), 160, 0.27),
]
PAY_LAUNCH = (2020, 6)
CLIENT_KINDS = [("Game publisher", "game"), ("Indie studio", "game"), ("Entertainment company", "game"),
                ("Regional bank", "fin"), ("Credit union", "fin"), ("Fintech startup", "fin"),
                ("Retailer", "fin"), ("Insurance company", "fin")]
PROJECT_TYPES = {"game": ["Mobile game development", "PC game port", "Web game", "Multiplayer backend",
                          "Live-ops support retainer", "Game art & animation", "Game analytics platform"],
                 "fin": ["Mobile banking app", "Payment gateway integration", "Fraud analytics dashboard",
                         "KYC onboarding flow", "Core ledger modernisation", "Open banking API", "Card issuing program"]}
CLIENT_A = """Atlas Brightline Copperleaf Northgate Silverpine Bluewater Redstone Ironbridge Clearwater Summit Pinecrest
Harborview Lakeshore Stonefield Oakridge Westbrook Eastwind Riverbend Highland Granite Meridian Keystone Sterling
Evergreen Cobalt Falcon Orion Pioneer Vantage Beacon Crescent Horizon Liberty Monarch Sapphire Willow""".split()
CLIENT_B = {"Game publisher": ["Games", "Interactive", "Entertainment", "Publishing"], "Indie studio": ["Studios", "Games Co."],
            "Entertainment company": ["Media Group", "Entertainment"], "Regional bank": ["Bank", "Savings Bank", "Bancorp"],
            "Credit union": ["Credit Union", "Federal Credit Union"], "Fintech startup": ["Pay", "Financial", "Money"],
            "Retailer": ["Stores", "Retail Group", "Outfitters"], "Insurance company": ["Insurance", "Mutual"]}
MERCHANT_CATS = [("Online retail", 5999, 62), ("Digital goods & games", 5816, 48), ("Subscriptions", 5968, 25),
                 ("Restaurants", 5812, 31), ("Travel", 4722, 280), ("Education", 8299, 90), ("Health & fitness", 7997, 55),
                 ("Marketplaces", 5262, 74)]


def ym_add(ym, n):
    t = ym[0] * 12 + ym[1] - 1 + n
    return (t // 12, t % 12 + 1)


def ym_s(ym):
    return f"{ym[0]}-{ym[1]:02d}"


def mdays(ym):
    nxt = ym_add(ym, 1)
    return (dt.date(nxt[0], nxt[1], 1) - dt.date(ym[0], ym[1], 1)).days


def months_between(a, b):
    return (b[0] - a[0]) * 12 + b[1] - a[1]


def pois(rng, mu):
    if mu <= 0:
        return 0
    if mu > 30:
        return max(0, int(round(rng.gauss(mu, math.sqrt(mu)))))
    L, k, p = math.exp(-mu), 0, 1.0
    while True:
        p *= rng.random()
        if p < L:
            return k
        k += 1


def r2(x):
    return round(x + 1e-9, 2)


def unique(rng, used, make, city):
    """A name not used yet: a few random tries, then the city, then a number."""
    for _ in range(12):
        name = make()
        if name not in used:
            break
    else:
        name = f"{make()} {city}"
        k = 2
        while name in used:
            name = f"{make()} {city} {k}"
            k += 1
    used.add(name)
    return name


# ---------------------------------------------------------------- documents

def new_doc(title, subtitle=None):
    d = Document()
    d.styles["Normal"].font.name = "Calibri"
    d.styles["Normal"].font.size = Pt(10.5)
    d.core_properties.author = CO
    d.core_properties.comments = BANNER
    sec = d.sections[0]
    p = sec.header.paragraphs[0]
    p.text = f"{CO} | Confidential (fictional)"
    p.runs[0].font.size, p.runs[0].font.color.rgb = Pt(8), RGBColor(0x55, 0x55, 0x55)
    f = sec.footer.paragraphs[0]
    f.text = BANNER
    f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    f.runs[0].font.size = Pt(7)
    d.add_heading(title, 0)
    if subtitle:
        d.add_paragraph(subtitle).runs[0].italic = True
    return d


def doc_table(d, header, rows):
    t = d.add_table(rows=1, cols=len(header))
    t.style = "Light Grid Accent 1"
    for i, h in enumerate(header):
        t.rows[0].cells[i].text = str(h)
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = "" if v is None else (f"{v:,.2f}" if isinstance(v, float) else str(v))
    return t


def save_doc(sink, rel, d):
    p = sink.path(rel)
    d.save(p)
    sink.wrote(p)


def bullet(d, text):
    d.add_paragraph(text, style="List Bullet")


def gz_csv(sink, rel, header, rows_iter):
    p = sink.path(rel)
    with gzip.open(p, "wt", encoding="utf-8", compresslevel=6, newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        for row in rows_iter:
            w.writerow(row)
    sink.wrote(p)


# ---------------------------------------------------------------- world

class World:
    def __init__(self, end, scale=1.0, seed=SEED):
        rng = random.Random(seed)
        self.scale, self.end = scale, end
        self.clients, self.projects, self.staff, self.merchants = [], [], [], []
        used, people = set(), set()

        def person():
            while True:
                n = f"{rng.choice(N.FIRST)} {rng.choice(N.LAST)}"
                if n not in people:
                    people.add(n)
                    return n

        self.person = person
        ym, i = START, 0
        while ym <= end:
            # new clients and projects
            for _ in range(pois(rng, (3 if i == 0 else 0.55 + 0.004 * i) * scale)):
                self._client(rng, ym, used, person)
            for c in self.clients:
                if c["end"] is None and c["since"] < ym and rng.random() < 0.05:
                    self._project(rng, c, ym)
                if c["end"] is None and c["since"] < ym and rng.random() < 0.006:
                    c["end"] = ym
            # merchants on CorvalisPay
            if ym >= PAY_LAUNCH:
                k = months_between(PAY_LAUNCH, ym)
                for _ in range(pois(rng, (6 + 0.9 * k) * scale)):
                    self._merchant(rng, ym, used)
                for mm in self.merchants:
                    if mm["end"] is None and mm["since"] < ym and rng.random() < 0.012:
                        mm["end"] = ym
            # staffing
            p, g, m = self.load(ym)
            for e in self.staff:
                if e["left"] is None and e["role"] not in HEADS.values() and rng.random() < 0.012:
                    e["left"] = ym
                    e["leave_reason"] = rng.choice(["Resigned", "Resigned", "Resigned", "Relocated", "Laid off",
                                                    "Dismissed", "Contract ended"])
            for role, (dept, base, rule) in ROLES.items():
                need = rule(p, g, m)
                have = sum(1 for e in self.staff if e["role"] == role and e["left"] is None)
                for _ in range(max(0, need - have)):
                    self._hire(rng, role, dept, base, ym, person)
            ym, i = ym_add(ym, 1), i + 1
        for e in self.staff:
            e["manager"] = self._manager(e)

    def _client(self, rng, ym, used, person):
        kind, line = rng.choice(CLIENT_KINDS)
        city, country, tax = N.location(rng)
        name = unique(rng, used, lambda: f"{rng.choice(CLIENT_A)} {rng.choice(CLIENT_B[kind])}", city)
        c = {"id": len(self.clients), "code": f"C{len(self.clients) + 1:04d}", "name": name, "kind": kind, "line": line,
             "city": city, "country": country, "tax": tax, "since": ym, "end": None, "contact": person(),
             "domain": safe(name.split()[0].lower(), 20) + ".example", "terms": rng.choice([30, 30, 45, 60])}
        self.clients.append(c)
        self._project(rng, c, ym)

    def _project(self, rng, c, ym):
        ptype = rng.choice(PROJECT_TYPES[c["line"]])
        months = rng.randint(3, 14) if "retainer" not in ptype else rng.randint(6, 24)
        team = rng.randint(3, 9)
        model = "Retainer" if "retainer" in ptype else rng.choice(["Fixed price", "Time & materials", "Time & materials"])
        rate = r2(rng.uniform(115, 185))
        hours = team * 140 * months
        p = {"id": len(self.projects), "code": f"PRJ-{ym[0]}-{len(self.projects) + 1:04d}", "client": c["id"],
             "name": f"{c['name'].split()[0]} {ptype}", "type": ptype, "line": c["line"], "start": ym_add(ym, 1),
             "end": ym_add(ym, months), "team": team, "model": model, "rate": rate,
             "budget": r2(hours * rate * rng.uniform(0.85, 1.1)), "risk": rng.choice(["Low", "Low", "Medium", "High"])}
        self.projects.append(p)

    def _merchant(self, rng, ym, used):
        cat, mcc, ticket = rng.choice(MERCHANT_CATS)
        city, country, _ = N.location(rng)
        name = unique(rng, used, lambda: f"{rng.choice(CLIENT_A + N.HOTEL_A)} "
                                         f"{rng.choice(['Goods', 'Market', 'Labs', 'Co.', 'Shop', 'Studio', 'Club', 'Hub', 'Supply', 'Outfitters', 'Store', 'Works'])}",
                      city)
        mid = len(self.merchants)
        self.merchants.append({"id": mid, "code": f"M{mid + 1:06d}", "name": name, "category": cat, "mcc": mcc,
                               "ticket": ticket * rng.uniform(0.6, 1.6), "volume": rng.lognormvariate(5.2, 1.1),
                               "city": city, "country": country, "since": ym, "end": None,
                               "risk": rng.choice(["Low", "Low", "Low", "Medium", "High"]),
                               "rate": rng.choice([0.029, 0.027, 0.025, 0.024]), "fixed": 0.30})

    def _hire(self, rng, role, dept, base, ym, person):
        n = person()
        first, last = n.split(" ", 1)
        self.staff.append({"id": f"E{len(self.staff) + 1:05d}", "name": n,
                           "email": f"{first.lower()}.{last.lower().replace(' ', '').replace('-', '')}@{DOMAIN}",
                           "role": role, "dept": dept, "hired": ym, "left": None, "leave_reason": "",
                           "base": r2(base * rng.uniform(0.9, 1.12) * 1.035 ** (ym[0] - 2018)),
                           "nationality": rng.choice(N.NATIONALITIES), "crew": len(self.staff) % 4,
                           "level": rng.choice(["I", "II", "II", "III", "Senior", "Staff"])})

    def _manager(self, e):
        head = HEADS.get(e["dept"])
        if e["role"] == "Chief Executive Officer":
            return ""
        if e["role"] == head or e["role"] in HEADS.values():
            return self.staff[0]["id"]
        return next((x["id"] for x in self.staff if x["role"] == head), self.staff[0]["id"])

    # queries
    def active_projects(self, ym):
        live = {c["id"] for c in self.clients if c["end"] is None or c["end"] > ym}
        return [p for p in self.projects if p["start"] <= ym <= p["end"] and p["client"] in live]

    def games_live(self, ym):
        return [g for g in GAMES if g[3] <= ym]

    def dau(self, game, ym, day=0):
        name, genre, plat, launch, peak, arpdau = game
        age = months_between(launch, ym) + day / 30
        if age < 0:
            return 0
        curve = (1 - math.exp(-age / 4)) * math.exp(-max(age - 18, 0) / 60)
        return peak * 1000 * curve * self.scale

    def active_merchants(self, ym):
        return [m for m in self.merchants if m["since"] < ym and (m["end"] is None or m["end"] > ym)]

    def load(self, ym):
        players = sum(self.dau(g, ym) for g in self.games_live(ym)) / 1000
        return len(self.active_projects(ym)), int(players), len(self.active_merchants(ym))

    def staff_at(self, ym, role=None):
        return [e for e in self.staff if e["hired"] <= ym and (e["left"] is None or e["left"] > ym)
                and (role is None or e["role"] == role)]

    def pay(self, e, ym):
        return r2(e["base"] * 1.035 ** (ym[0] - e["hired"][0]))


# ---------------------------------------------------------------- month

class Month:
    def __init__(self, w, ym, ledger, sink, detail=1.0):
        self.w, self.ym, self.L, self.sink, self.detail = w, ym, ledger, sink, detail
        self.rng = random.Random(f"{SEED}-{ym_s(ym)}")
        self.tag, self.y = ym_s(ym), str(ym[0])
        self.days = mdays(ym)
        self.d0 = dt.date(ym[0], ym[1], 1)
        self.eom = self.d0 + dt.timedelta(days=self.days - 1)
        self.staff = w.staff_at(ym)
        self.projects = w.active_projects(ym)
        self.cl_by = {c["id"]: c for c in w.clients}
        self.C, self.F, self.BANNER = sys.modules[__name__], F, BANNER
        self.gl, self.emails = [], []
        self.work_counts = {}
        self.revenue = {}

    def people(self, *roles):
        return [e for e in self.staff if e["role"] in roles] or self.staff[:1]

    def ts(self, day=None, hour=None):
        r = self.rng
        d = self.d0 + dt.timedelta(days=r.randrange(self.days) if day is None else day)
        return dt.datetime.combine(d, dt.time(r.randint(8, 19) if hour is None else hour, r.randrange(60), r.randrange(60)))

    def count(self, eid, k, n=1):
        self.work_counts.setdefault(eid, [0, 0, 0])[k] += n

    def post(self, date, ref, desc, lines):
        assert abs(sum(d - c for _, d, c in lines)) < 0.01, (ref, lines)
        for acct, dr, cr in lines:
            if dr or cr:
                self.gl.append([date, ref, acct, ACCOUNTS[acct][0], desc, r2(dr), r2(cr)])

    def mail(self, t, frm, to, subject, body):
        m = EmailMessage()
        m["From"], m["To"], m["Subject"] = frm, to, subject
        m["Date"] = format_datetime(t.replace(tzinfo=dt.timezone.utc))
        m["Message-ID"] = f"<{self.tag}.{len(self.emails) + 1}@{DOMAIN}>"
        m["X-Dataset-Notice"] = BANNER
        m.set_content(body)
        self.emails.append((t, subject, m.as_bytes()))

    def run(self):
        self.delivery()
        self.engineering()
        self.games()
        self.payments()
        self.support()
        self.marketing()
        self.hr()
        self.finance()
        S.collect(self)
        if self.ym[1] == 12:
            S.person_year(self.w, self.L, self.ym[0], self.sink, self.C, final=True)
        self.reports()
        self.meetings()
        self.write_mail()

    # ---------------- delivery: projects, SOWs, status reports, timesheets
    def delivery(self):
        r = self.rng
        pms = self.people("Project Manager")
        devs = self.people("Backend Engineer", "Frontend Engineer", "Mobile Engineer", "Game Developer", "QA Engineer",
                           "Game Designer", "Technical Artist", "UX Designer", "DevOps Engineer")
        self.billable = {}
        ts_rows = []
        self.team = {}
        for p in self.projects:
            pr = random.Random(f"{SEED}-team-{p['id']}")
            team = pr.sample(devs, min(len(devs), p["team"]))
            self.team[p["id"]] = team
            hours = 0.0
            for e in team:
                h = r2(max(0, r.gauss(118, 22)))
                hours += h
                ts_rows.append([self.tag, p["code"], p["name"], e["id"], e["name"], e["role"], h, "Yes"])
            amount = r2(hours * p["rate"]) if p["model"] == "Time & materials" else r2(
                p["budget"] / max(1, months_between(p["start"], p["end"]) + 1))
            if p["model"] == "Retainer":
                amount = r2(p["budget"] / max(1, months_between(p["start"], p["end"]) + 1))
            self.billable[p["id"]] = (amount, r2(hours))
            if p["start"] == self.ym:
                self.sow(p, pms[p["id"] % len(pms)])
                if not any(q["client"] == p["client"] and q["start"] < p["start"] for q in self.w.projects):
                    self.msa(self.cl_by[p["client"]])
            self.status_report(p, pms[p["id"] % len(pms)], hours)
        wb = new_book()
        add_sheet(wb, "Timesheets", ["Month", "Project", "Project name", "Employee", "Name", "Role", "Hours", "Billable"],
                  ts_rows, {"Project name": 34, "Name": 22, "Role": 22})
        save_book(self.sink, f"{F['proj']}/{self.y}/Timesheets/Project timesheets {self.tag}.xlsx", wb, "Timesheets", BANNER)

    def msa(self, c):
        d = new_doc(f"Master Services Agreement - {c['name']}", f"{CO} and {c['name']} | effective {self.d0:%d %B %Y}")
        for h, t in [("1. Services", "Provider performs software development, game development, payments engineering "
                                     "and related services described in Statements of Work."),
                     ("2. Fees", f"As set out in each Statement of Work. Invoices are payable within {c['terms']} days."),
                     ("3. Intellectual property", "On full payment, Client owns the deliverables created for it. Provider keeps "
                                                 "its pre-existing tools and libraries and grants Client a perpetual licence to them."),
                     ("4. Confidentiality", "Each party protects the other's confidential information for five years."),
                     ("5. Data protection", "Where Provider processes personal data for Client, the Data Processing "
                                            "Addendum applies. Payment card data is handled only in PCI DSS compliant systems."),
                     ("6. Warranties", "Deliverables materially conform to their specification for 90 days after acceptance."),
                     ("7. Liability", "Each party's liability is capped at the fees paid in the 12 months before the claim."),
                     ("8. Term", "Three years, renewing yearly unless either party gives 60 days' notice.")]:
            d.add_heading(h, 1)
            d.add_paragraph(t)
        doc_table(d, ["For", "Name", "Title"], [[CO, self.w.staff[0]["name"], "Chief Executive Officer"],
                                               [c["name"], c["contact"], "Authorised signatory"]])
        save_doc(self.sink, f"{F['legal']}/Client contracts/{safe(c['name'], 60)}/MSA {c['name']}.docx", d)

    def sow(self, p, pm):
        c = self.cl_by[p["client"]]
        months = months_between(p["start"], p["end"]) + 1
        d = new_doc(f"Statement of Work {p['code']}", f"{p['name']} for {c['name']}")
        doc_table(d, ["Item", "Detail"], [["Project", p["name"]], ["Type", p["type"]], ["Pricing", p["model"]],
                                         ["Start", ym_s(p["start"])], ["End", ym_s(p["end"])], ["Team size", p["team"]],
                                         ["Blended rate (USD/h)", p["rate"]], ["Budget (USD)", p["budget"]],
                                         ["Project manager", pm["name"]]])
        d.add_heading("Scope", 1)
        for x in self.rng.sample(["Discovery workshops and backlog", "UX flows and visual design", "Core build in two-week sprints",
                                  "Automated test suite and CI pipeline", "Security review and penetration test",
                                  "Load testing to target concurrency", "App store / platform submission",
                                  "Production launch and hypercare", "Analytics and dashboards", "Knowledge transfer"], 6):
            bullet(d, x)
        d.add_heading("Milestones", 1)
        doc_table(d, ["Milestone", "Month", "Payment %"],
                  [[m, ym_s(ym_add(p["start"], int((months - 1) * f))), pc] for m, f, pc in
                   [("Discovery complete", 0.1, 15), ("Alpha", 0.45, 25), ("Beta", 0.75, 30), ("Launch", 1.0, 30)]])
        d.add_heading("Acceptance", 1)
        d.add_paragraph("Each milestone is accepted when it passes the agreed test plan, or 10 business days after "
                        "delivery without a written rejection.")
        save_doc(self.sink, f"{F['legal']}/Client contracts/{safe(c['name'], 60)}/{p['code']} Statement of Work.docx", d)

    def status_report(self, p, pm, hours):
        r = self.rng
        c = self.cl_by[p["client"]]
        done = months_between(p["start"], self.ym) + 1
        total = months_between(p["start"], p["end"]) + 1
        rag = r.choices(["Green", "Amber", "Red"], [70, 22, 8] if p["risk"] != "High" else [45, 38, 17])[0]
        d = new_doc(f"Project status - {p['name']}", f"{c['name']} | {self.d0:%B %Y} | {p['code']}")
        doc_table(d, ["Measure", "Value"], [["Overall status", rag], ["Month", f"{done} of {total}"],
                                            ["Hours this month", hours], ["Budget used %", round(100 * done / total)],
                                            ["Open risks", r.randint(0, 6)], ["Change requests", r.randint(0, 3)]])
        d.add_heading("Done this month", 1)
        for x in r.sample(["Completed sprint goals", "Integrated payment provider sandbox", "Shipped build to client QA",
                           "Fixed crash on older Android devices", "Finished level art for chapter 2",
                           "Load-tested matchmaking to 50k CCU", "Passed internal security review",
                           "Implemented KYC document upload", "Added analytics events", "Refactored ledger posting service"], 3):
            bullet(d, x)
        d.add_heading("Risks and issues", 1)
        for x in r.sample(["Client API documentation incomplete", "Key engineer on leave next month",
                           "App store review delay", "Scope creep on reporting features", "Third-party SDK licence cost",
                           "Performance on low-end devices"], 2):
            bullet(d, x)
        save_doc(self.sink, f"{F['proj']}/{self.y}/Status reports/{self.tag}/{safe(p['name'], 60)} {self.tag}.docx", d)

    # ---------------- engineering: issues, commits, pull requests, builds, incidents
    def engineering(self):
        r, k = self.rng, self.detail
        eng = self.people("Backend Engineer", "Frontend Engineer", "Mobile Engineer", "Game Developer", "QA Engineer",
                          "DevOps Engineer", "Security Engineer", "Data Analyst", "Technical Artist")
        repos = ["corvalis-platform", "internal-tools", "infrastructure"] + \
                [f"{safe(p['name'], 30).lower().replace(' ', '-')}" for p in self.projects] + \
                [g[0].lower().replace(" ", "-") for g in self.w.games_live(self.ym)] + \
                (["corvalispay-api", "corvalispay-ledger", "corvalispay-risk", "merchant-dashboard"]
                 if self.ym >= PAY_LAUNCH else [])
        issues, commits, prs, builds = [], [], [], []
        types = ["Story", "Story", "Task", "Bug", "Bug", "Spike", "Epic"]
        msgs = ["Fix {x} on {y}", "Add {x} to {y}", "Refactor {y} {x}", "Update {x} handling", "Improve {y} performance",
                "Handle edge case in {x}", "Bump dependencies for {y}", "Add tests for {x}", "Remove dead code in {y}"]
        things = ["retry logic", "currency rounding", "session timeout", "leaderboard", "inventory sync", "webhook",
                  "refund flow", "matchmaking", "push notifications", "login", "KYC upload", "chargeback export",
                  "asset loading", "save games", "rate limiter", "audit log", "pagination", "localisation"]
        for repo in repos:
            for _ in range(pois(r, 60 * k)):
                e = r.choice(eng)
                t = self.ts()
                typ = r.choice(types)
                status = r.choices(["Done", "Done", "In Progress", "In Review", "To Do", "Won't Do"], [50, 10, 15, 10, 12, 3])[0]
                resolved = (t + dt.timedelta(hours=r.lognormvariate(3.5, 1))) if status == "Done" else None
                key = f"{repo[:6].upper().replace('-', '')}-{r.randint(100, 9999)}"
                issues.append([key, repo, typ, r.choice(["Highest", "High", "Medium", "Medium", "Low"]), status,
                               e["name"], r.choice(self.staff)["name"], t.strftime("%Y-%m-%d %H:%M"),
                               resolved.strftime("%Y-%m-%d %H:%M") if resolved else "", r.choice([1, 2, 3, 5, 8, 13]),
                               f"Sprint {self.ym[0] % 100}.{(self.ym[1] - 1) * 2 + r.randint(1, 2)}",
                               f"{r.choice(msgs).format(x=r.choice(things), y=repo)}"])
                if status == "Done":
                    self.count(e["id"], 1)
            for _ in range(pois(r, 140 * k)):
                e = r.choice(eng)
                t = self.ts()
                adds, dels = int(r.lognormvariate(3.5, 1.2)), int(r.lognormvariate(2.8, 1.2))
                commits.append(["%040x" % r.getrandbits(160), repo, e["name"], e["email"], t.isoformat(),
                                r.choice(msgs).format(x=r.choice(things), y=repo), r.randint(1, 14), adds, dels])
                self.count(e["id"], 0)
            for _ in range(pois(r, 30 * k)):
                e = r.choice(eng)
                t = self.ts()
                revs = r.sample(eng, min(2, len(eng)))
                merged = t + dt.timedelta(hours=r.lognormvariate(2.6, 1))
                prs.append([r.randint(1, 9000), repo, e["name"], ", ".join(x["name"] for x in revs), t.strftime("%Y-%m-%d %H:%M"),
                            merged.strftime("%Y-%m-%d %H:%M"), r.randint(0, 25), int(r.lognormvariate(4.5, 1)),
                            r.choice(["Merged", "Merged", "Merged", "Closed"])])
            for _ in range(pois(r, 90 * k)):
                t = self.ts()
                st = r.choices(["success", "failed", "cancelled"], [84, 13, 3])[0]
                builds.append([r.randint(10 ** 5, 10 ** 6), repo, r.choice(["main", "main", "develop", "feature/x", "release"]),
                               st, t.strftime("%Y-%m-%d %H:%M:%S"), int(r.lognormvariate(5.8, 0.5)),
                               r.choice(["unit", "integration", "e2e", "build", "deploy-staging", "deploy-prod"])])
        base = f"{F['eng']}/{self.y}/{self.tag}"
        gz_csv(self.sink, f"{base}/Issues export {self.tag}.csv.gz",
               ["Key", "Repository", "Type", "Priority", "Status", "Assignee", "Reporter", "Created", "Resolved",
                "Story points", "Sprint", "Summary"], issues)
        gz_csv(self.sink, f"{base}/Git commit log {self.tag}.csv.gz",
               ["Commit", "Repository", "Author", "Email", "Date", "Message", "Files changed", "Additions", "Deletions"], commits)
        wb = new_book()
        add_sheet(wb, "Pull requests", ["PR", "Repository", "Author", "Reviewers", "Opened", "Merged/closed", "Comments",
                                        "Lines changed", "Outcome"], prs, {"Reviewers": 40, "Repository": 28})
        save_book(self.sink, f"{base}/Pull requests {self.tag}.xlsx", wb, "Pull requests", BANNER)
        gz_csv(self.sink, f"{base}/CI builds {self.tag}.csv.gz",
               ["Build", "Repository", "Branch", "Status", "Started", "Duration seconds", "Stage"], builds)
        # production incidents
        oncall = self.people("DevOps Engineer", "Backend Engineer", "Security Engineer")
        incs = []
        for _ in range(pois(r, 2.5 + len(repos) / 6)):
            t = self.ts(hour=r.randrange(24))
            sev = r.choices(["SEV1", "SEV2", "SEV3", "SEV4"], [5, 20, 45, 30])[0]
            svc = r.choice(repos)
            dur = int(r.lognormvariate(3.6 if sev in ("SEV1", "SEV2") else 3, 0.8))
            cause = r.choice(["Bad deploy", "Database connection pool exhausted", "Expired TLS certificate",
                              "Third-party payment provider outage", "Cache stampede after release", "DDoS on login",
                              "Config change without review", "Disk full on log volume", "Memory leak in game server"])
            iid = f"INC-{self.tag.replace('-', '')}-{len(incs) + 1:03d}"
            incs.append([iid, sev, svc, t.strftime("%Y-%m-%d %H:%M"), dur, cause, r.choice(oncall)["name"],
                         "Yes" if sev in ("SEV1", "SEV2") else "No"])
            if sev in ("SEV1", "SEV2"):
                d = new_doc(f"Postmortem {iid}: {svc}", f"{sev} | {t:%d %b %Y %H:%M} UTC | {dur} minutes")
                d.add_heading("Summary", 1)
                d.add_paragraph(f"{svc} was degraded for {dur} minutes. Root cause: {cause.lower()}.")
                d.add_heading("Timeline", 1)
                for m_, txt in [(0, "Alert fired from SLO burn-rate monitor"), (3, "On-call engineer acknowledged"),
                                (12, "Incident channel opened, status page updated"), (dur - 5, "Mitigation deployed"),
                                (dur, "Service recovered; monitoring")]:
                    bullet(d, f"{(t + dt.timedelta(minutes=max(m_, 0))):%H:%M} - {txt}")
                d.add_heading("What went well / what did not", 1)
                bullet(d, "Alerting caught the problem before customers reported it.")
                bullet(d, "Runbook for this service was out of date.")
                d.add_heading("Action items", 1)
                doc_table(d, ["Action", "Owner", "Due"], [[a, r.choice(oncall)["name"], (self.eom + dt.timedelta(days=14)).isoformat()]
                                                          for a in r.sample(["Add canary stage to deploys", "Alert on certificate expiry",
                                                                             "Raise connection pool and add backpressure",
                                                                             "Update runbook", "Add load test to CI",
                                                                             "Second payment provider failover"], 3)])
                save_doc(self.sink, f"{F['eng']}/{self.y}/Postmortems/{iid} {safe(svc, 40)}.docx", d)
        self.incidents = incs
        wb = new_book()
        add_sheet(wb, "Incidents", ["Incident", "Severity", "Service", "Started", "Minutes", "Cause", "On call", "Postmortem"],
                  incs, {"Cause": 40, "Service": 28})
        save_book(self.sink, f"{base}/Production incidents {self.tag}.xlsx", wb, "Incidents", BANNER)
        self.repos = repos

    # ---------------- live games: KPIs, purchases, telemetry
    def games(self):
        r, k = self.rng, self.detail
        self.game_rev, self.ad_rev = 0.0, 0.0
        for g in self.w.games_live(self.ym):
            name, genre, plat, launch, peak, arpdau = g
            base = f"{F['games']}/{safe(name, 40)}/{self.y}"
            kpis, purchases = [], []
            installs_total = 0
            for d in range(self.days):
                day = self.d0 + dt.timedelta(days=d)
                dau = int(self.w.dau(g, self.ym, d) * r.uniform(0.92, 1.08) * (1.12 if day.weekday() >= 5 else 1))
                if dau <= 0:
                    continue
                installs = int(dau * r.uniform(0.02, 0.06))
                installs_total += installs
                rev = dau * arpdau * r.uniform(0.85, 1.15)
                ads = dau * 0.012 * r.uniform(0.8, 1.2)
                payers = int(dau * r.uniform(0.012, 0.03))
                self.game_rev += rev * 0.7  # net of store fees
                self.ad_rev += ads
                kpis.append([day.isoformat(), dau, installs, int(dau * r.uniform(2.6, 3.4)),
                             round(r.uniform(0.32, 0.45), 3), round(r.uniform(0.1, 0.18), 3), round(r.uniform(0.03, 0.07), 3),
                             payers, r2(rev), r2(ads), r2(rev / dau), round(r.uniform(0.2, 1.2), 2)])
                for _ in range(min(payers, int(400 * k))):
                    sku, price = r.choice([("gems_small", 0.99), ("gems_medium", 4.99), ("gems_large", 19.99),
                                           ("battle_pass", 9.99), ("starter_pack", 2.99), ("vip_month", 14.99)])
                    purchases.append([f"T{r.getrandbits(48):012x}", day.isoformat(), f"P{r.randrange(10 ** 8):08d}", sku,
                                      price, r.choice(["App Store", "Google Play", "Steam", "Web"]),
                                      N.location(r)[1], "Yes" if r.random() < 0.006 else "No"])
            wb = new_book()
            add_sheet(wb, "Daily KPIs", ["Date", "DAU", "Installs", "Sessions", "D1 retention", "D7 retention", "D30 retention",
                                         "Payers", "IAP revenue USD", "Ad revenue USD", "ARPDAU", "Crash rate %"], kpis)
            save_book(self.sink, f"{base}/KPIs/{safe(name, 40)} daily KPIs {self.tag}.xlsx", wb, "KPIs", BANNER)
            gz_csv(self.sink, f"{base}/Purchases/{safe(name, 40)} purchases {self.tag}.csv.gz",
                   ["Transaction", "Date", "Player", "SKU", "Price USD", "Store", "Country", "Refunded"], purchases)
            if k > 0:
                self.telemetry(g, base)
            if r.random() < 0.6:
                d = new_doc(f"{name} - patch notes {self.tag}")
                for x in r.sample(["New event: Harvest Festival", "Balance changes to three heroes", "Fixed rare crash on load",
                                   "Matchmaking improvements", "New cosmetic shop items", "Reduced download size",
                                   "Added two new levels", "Daily rewards rework"], 4):
                    bullet(d, x)
                save_doc(self.sink, f"{base}/Patch notes/{safe(name, 40)} patch notes {self.tag}.docx", d)

    def telemetry(self, g, base):
        r, k = self.rng, self.detail
        name = g[0]
        events = ["session_start", "level_start", "level_complete", "level_fail", "store_open", "purchase",
                  "ad_impression", "tutorial_step", "friend_invite", "session_end", "crash"]
        wts = [10, 14, 10, 5, 4, 0.6, 9, 2, 0.5, 10, 0.2]
        sample = int(min(30000, self.w.dau(g, self.ym) * 0.02) * k)  # 2% of players, capped

        def rows():
            players = [f"P{r.randrange(10 ** 8):08d}" for _ in range(max(sample, 1))]
            for p in players:
                for _ in range(pois(r, 6)):
                    day = r.randrange(self.days)
                    t = dt.datetime.combine(self.d0 + dt.timedelta(days=day), dt.time()) + dt.timedelta(seconds=r.randrange(86400))
                    for ev in r.choices(events, wts, k=r.randint(3, 12)):
                        t += dt.timedelta(seconds=r.randint(5, 240))
                        yield [t.strftime("%Y-%m-%dT%H:%M:%SZ"), p, ev, r.randint(1, 120),
                               r.choice(["ios", "android", "windows", "web"]), r.choice(["1.%d.%d" % (r.randint(0, 9), r.randint(0, 20))])]
        gz_csv(self.sink, f"{base}/Telemetry/{safe(name, 40)} events sample {self.tag}.csv.gz",
               ["Timestamp", "Player", "Event", "Level", "Platform", "Build"], rows())

    # ---------------- CorvalisPay: merchants, transactions, disputes, KYC, settlement
    def payments(self):
        r, k = self.rng, self.detail
        self.pay_rev = self.pay_cost = self.fraud_loss = 0.0
        if self.ym < PAY_LAUNCH:
            return
        merchants = self.w.active_merchants(self.ym)
        base = f"{F['pay']}/{self.y}/{self.tag}"
        summary, disputes, alerts = [], [], []

        per_merchant = int(1200000 * k / max(len(merchants), 1))  # rows written per month stay bounded

        def tx_rows():
            for m in merchants:
                n_real = int(m["volume"] * (1 + 0.02 * months_between(m["since"], self.ym)))
                n = min(n_real, per_merchant, int(3000 * k))
                vol = fees = 0.0
                for _ in range(n):
                    t = dt.datetime.combine(self.d0, dt.time()) + dt.timedelta(seconds=r.randrange(self.days * 86400))
                    amt = r2(max(1, r.lognormvariate(math.log(m["ticket"]), 0.7)))
                    method = r.choices(["card", "card", "card", "ach", "wallet"])[0]
                    risk = round(r.betavariate(1.2, 9) * 100, 1)
                    status = "declined" if r.random() < 0.06 + risk / 400 else "captured"
                    fraud = status == "captured" and risk > 70 and r.random() < 0.25
                    fee = r2(amt * m["rate"] + m["fixed"]) if status == "captured" else 0.0
                    if status == "captured":
                        vol += amt
                        fees += fee
                    yield [f"tx_{r.getrandbits(64):016x}", t.strftime("%Y-%m-%dT%H:%M:%SZ"), m["code"], amt, "USD", method,
                           r.choice(["visa", "mastercard", "amex", "discover"]) if method == "card" else "",
                           N.location(r)[1], status, r.choice(["insufficient_funds", "do_not_honor", "suspected_fraud",
                                                               "expired_card"]) if status == "declined" else "",
                           risk, "Yes" if fraud else "No", fee]
                    if fraud:
                        alerts.append([self.tag, m["code"], m["name"], amt, risk, r.choice(["velocity", "geo mismatch",
                                                                                             "card testing", "BIN attack"]),
                                       r.choice(["Confirmed fraud", "False positive", "Under review"])])
                scale = n_real / max(n, 1)
                summary.append([m["code"], m["name"], m["category"], m["risk"], n_real, r2(vol * scale), r2(fees * scale)])
                for _ in range(pois(r, n_real * 0.0012)):
                    reason = r.choice(["Fraudulent", "Product not received", "Duplicate", "Subscription cancelled",
                                       "Credit not processed"])
                    amt = r2(r.lognormvariate(math.log(m["ticket"]), 0.6))
                    disputes.append([f"dp_{r.getrandbits(40):010x}", m["code"], m["name"], amt, reason,
                                     r.choice(["Won", "Lost", "Lost", "Pending"]), (self.d0 + dt.timedelta(days=r.randrange(self.days))).isoformat()])
        gz_csv(self.sink, f"{base}/Transactions {self.tag}.csv.gz",
               ["Transaction", "Created", "Merchant", "Amount", "Currency", "Method", "Card brand", "Country", "Status",
                "Decline reason", "Risk score", "Fraud flag", "Fee USD"], tx_rows())
        wb = new_book()
        add_sheet(wb, "Merchant volume", ["Merchant", "Name", "Category", "Risk tier", "Transactions", "Volume USD",
                                          "Fees USD"], summary, {"Name": 28})
        add_sheet(wb, "Disputes", ["Dispute", "Merchant", "Name", "Amount USD", "Reason", "Outcome", "Opened"], disputes,
                  {"Name": 28, "Reason": 24})
        add_sheet(wb, "Fraud alerts", ["Month", "Merchant", "Name", "Amount USD", "Risk score", "Rule", "Outcome"], alerts,
                  {"Name": 28})
        save_book(self.sink, f"{base}/Merchant volume, disputes & fraud {self.tag}.xlsx", wb, "Payments", BANNER)
        self.pay_rev = sum(x[6] for x in summary)
        gross = sum(x[5] for x in summary)
        self.pay_cost = r2(gross * 0.0175)  # interchange + network
        self.fraud_loss = r2(sum(x[3] for x in disputes if x[5] == "Lost"))
        # KYC / KYB for new merchants
        kyc = [[m["code"], m["name"], m["category"], m["country"], r.choice(["Passed", "Passed", "Passed", "Manual review", "Rejected"]),
                r.choice(["Business registration", "Beneficial owners", "Bank account", "Website review", "Sanctions screening"]),
                r.choice(self.people("Compliance Officer", "Payments Operations Analyst"))["name"]]
               for m in self.w.merchants if m["since"] == self.ym]
        wb = new_book()
        add_sheet(wb, "KYB checks", ["Merchant", "Name", "Category", "Country", "Result", "Check", "Reviewer"], kyc, {"Name": 28})
        save_book(self.sink, f"{base}/Merchant onboarding (KYB) {self.tag}.xlsx", wb, "KYB", BANNER)
        for m in self.w.merchants:
            if m["since"] == self.ym and r.random() < 0.25:
                d = new_doc(f"Merchant Services Agreement - {m['name']}", f"CorvalisPay | {m['code']}")
                doc_table(d, ["Term", "Value"], [["Category", m["category"]], ["MCC", m["mcc"]],
                                                ["Pricing", f"{m['rate'] * 100:.1f}% + ${m['fixed']:.2f} per transaction"],
                                                ["Settlement", "T+2 business days"], ["Reserve", "5% rolling 90 days" if m["risk"] == "High" else "None"],
                                                ["Risk tier", m["risk"]]])
                d.add_paragraph("The merchant follows card network rules, keeps refund and dispute evidence for 18 months, "
                                "and does not store full card numbers.")
                save_doc(self.sink, f"{F['legal']}/Merchant agreements/{self.y}/{m['code']} {safe(m['name'], 50)}.docx", d)
        self.disputes, self.fraud_alerts = disputes, alerts

    # ---------------- support
    def support(self):
        r = self.rng
        agents = self.people("Customer Support Agent")
        rows = []
        players = sum(self.w.dau(g, self.ym) for g in self.w.games_live(self.ym))
        n_player = pois(r, players / 400)
        n_merch = pois(r, len(self.w.active_merchants(self.ym)) * 0.8)
        cats_p = ["Purchase not received", "Account recovery", "Bug report", "Refund request", "Cheating report",
                  "Ban appeal", "Performance issue", "Feature request"]
        cats_m = ["Payout delayed", "Dispute help", "API integration", "Account verification", "Fee question",
                  "Refund processing", "Webhook failures", "Fraud question"]
        for kind, n, cats in (("Player", n_player, cats_p), ("Merchant", n_merch, cats_m)):
            for _ in range(n):
                a = r.choice(agents)
                t = self.ts(hour=r.randrange(24))
                fr = int(r.lognormvariate(4.2, 0.8))
                res = r2(r.lognormvariate(2.4, 1))
                rows.append([f"SUP-{r.randrange(10 ** 7):07d}", t.strftime("%Y-%m-%d %H:%M"), kind, r.choice(cats),
                             r.choice(["Email", "In-app", "Chat", "Web form"]), r.choice(["Low", "Normal", "Normal", "High", "Urgent"]),
                             a["name"], fr, res, r.choices([5, 4, 3, 2, 1, None], [38, 27, 12, 6, 4, 13])[0]])
                self.count(a["id"], 2)
        rows.sort(key=lambda x: x[1])
        wb = new_book()
        add_sheet(wb, "Tickets", ["Ticket", "Opened", "Customer type", "Category", "Channel", "Priority", "Agent",
                                  "First response min", "Resolution hours", "CSAT"], rows, {"Category": 26, "Agent": 22})
        save_book(self.sink, f"{F['support']}/{self.y}/Support tickets {self.tag}.xlsx", wb, "Support", BANNER)
        self.tickets_n = len(rows)

    # ---------------- marketing & sales
    def marketing(self):
        r = self.rng
        camps = []
        self.ua_spend = 0.0
        for g in self.w.games_live(self.ym):
            for ch in ("Meta Ads", "Google Ads", "TikTok", "Apple Search Ads", "Unity Ads", "Influencers"):
                spend = r2(self.w.dau(g, self.ym) * r.uniform(0.02, 0.08))
                if spend <= 0:
                    continue
                cpi = r.uniform(0.8, 4.5)
                inst = int(spend / cpi)
                self.ua_spend += spend
                camps.append([self.tag, g[0], ch, spend, inst, r2(spend / max(inst, 1)), round(r.uniform(0.05, 0.6), 2),
                              round(r.uniform(0.2, 1.6), 2)])
        wb = new_book()
        add_sheet(wb, "UA campaigns", ["Month", "Game", "Channel", "Spend USD", "Installs", "CPI", "D7 ROAS", "D90 ROAS"], camps)
        save_book(self.sink, f"{F['mkt']}/{self.y}/User acquisition {self.tag}.xlsx", wb, "UA", BANNER)
        ae = self.people("Account Executive")
        opp = []
        for c in self.w.clients:
            lead = months_between(self.ym, c["since"])
            if 0 <= lead <= 3:
                stage = "Closed won" if lead == 0 else ["Proposal", "Discovery", "Qualified"][min(lead - 1, 2)]
                opp.append([c["name"], c["kind"], c["city"], stage, PROJECT_TYPES[c["line"]][c["id"] % 7],
                            r2(r.uniform(80, 900) * 1000), {"Qualified": 20, "Discovery": 40, "Proposal": 65, "Closed won": 100}[stage],
                            ae[c["id"] % len(ae)]["name"]])
        for _ in range(pois(r, 6)):
            kind, line = r.choice(CLIENT_KINDS)
            st = r.choice(["Qualified", "Discovery", "Proposal", "Closed lost", "Closed lost"])
            opp.append([f"{r.choice(CLIENT_A)} {r.choice(CLIENT_B[kind])} (prospect)", kind, N.location(r)[0], st,
                        r.choice(PROJECT_TYPES[line]), r2(r.uniform(60, 700) * 1000),
                        0 if st == "Closed lost" else {"Qualified": 20, "Discovery": 40, "Proposal": 65}[st], r.choice(ae)["name"]])
        wb = new_book()
        add_sheet(wb, "Pipeline", ["Account", "Segment", "City", "Stage", "Opportunity", "Value USD", "Probability %", "Owner"],
                  opp, {"Account": 34, "Opportunity": 30})
        ws = wb["Pipeline"]
        n = ws.max_row
        ws.cell(n + 2, 5, "Weighted pipeline")
        ws.cell(n + 2, 6, f"=SUMPRODUCT(F2:F{n},G2:G{n})/100")
        save_book(self.sink, f"{F['crm']}/Sales pipeline/{self.y}/Opportunities {self.tag}.xlsx", wb, "Pipeline", BANNER)

    # ---------------- HR and payroll
    def hr(self):
        r = self.rng
        pay, self.pay_lines = [], {}
        self.payroll = self.taxes = 0.0
        for e in self.staff:
            base = self.w.pay(e, self.ym)
            leave = r.choices([0, 0, 0, 1, 2, 5], [45, 20, 10, 10, 8, 7])[0]
            bonus = r2(base * r.uniform(0.05, 0.15)) if self.ym[1] == 3 and r.random() < 0.7 else 0.0
            gross = r2(base + bonus)
            tax = r2(gross * 0.0765 + gross * 0.08)  # employer FICA + benefits
            self.payroll += gross
            self.taxes += tax
            self.pay_lines[e["id"]] = (base, 0.0, 0.0, 0, bonus, gross, tax, leave)
            pay.append([e["id"], e["name"], e["dept"], e["role"], e["level"], base, bonus, gross, r2(gross * 0.22),
                        r2(gross * 0.0765), r2(gross * 0.06), r2(gross * (1 - 0.22 - 0.0765 - 0.06)), tax])
        wb = new_book()
        add_sheet(wb, "Payroll", ["Employee", "Name", "Department", "Role", "Level", "Base", "Bonus", "Gross",
                                  "Federal & state withholding", "Employee FICA", "401(k) & benefits", "Net pay",
                                  "Employer taxes & benefits"], pay, {"Name": 22, "Role": 26})
        ws = wb["Payroll"]
        n = ws.max_row
        ws.cell(n + 1, 1, "TOTAL")
        for col in "FGHIJKLM":
            ws[f"{col}{n + 1}"] = f"=SUM({col}2:{col}{n})"
        save_book(self.sink, f"{F['hr']}/{self.y}/Payroll register {self.tag}.xlsx", wb, "Payroll", BANNER)
        self.payroll, self.taxes = r2(self.payroll), r2(self.taxes)
        for e in self.w.staff:
            if e["hired"] == self.ym:
                S.joiner_docs(self, e)
            elif e["left"] == self.ym:
                S.leaver_docs(self, e)

    # ---------------- finance: double-entry books
    def finance(self):
        r, L = self.rng, self.L
        eom = self.eom
        S.expenses(self)
        if self.ym == START:
            self.post(self.d0, "OPEN", "Seed funding", [("1000", 15000000, 0), ("3000", 0, 15000000)])
        if self.ym in ((2020, 3), (2022, 6)):
            amt = 25000000 if self.ym == (2020, 3) else 60000000
            self.post(self.d0, f"RAISE-{self.tag}", "Equity financing round", [("1000", amt, 0), ("3000", 0, amt)])
        invoices = []
        for p in self.projects:
            amount, hours = self.billable[p["id"]]
            c = self.cl_by[p["client"]]
            tax = r2(amount * c["tax"] * 0.0)  # services mostly exempt from sales tax
            L["inv_no"] = L.get("inv_no", 0) + 1
            no = f"INV-{L['inv_no']:06d}"
            due = eom + dt.timedelta(days=c["terms"])
            delay = r.choices([0, 7, 20, 45, 90, None], [40, 25, 18, 10, 6, 1])[0]
            paid = None if delay is None else due + dt.timedelta(days=delay)
            invoices.append([no, eom.isoformat(), c["code"], c["name"], p["code"], p["model"], hours, amount, due.isoformat()])
            self.post(eom, no, f"Invoice {c['name']} {p['code']}", [("1100", amount, 0), ("4000", 0, amount)])
            L.setdefault("ar", []).append({"no": no, "client": c["name"], "due": due.isoformat(), "total": amount,
                                           "paid": paid.isoformat() if paid else None})
            self.invoice_pdf(no, eom, due, c, p, hours, amount)
        still = []
        for inv in L.get("ar", []):
            if inv["paid"] and inv["paid"] <= eom.isoformat():
                self.post(dt.date.fromisoformat(max(inv["paid"], self.d0.isoformat())), "RCPT " + inv["no"],
                          f"Receipt {inv['client']}", [("1000", inv["total"], 0), ("1100", 0, inv["total"])])
            elif not inv["paid"] and dt.date.fromisoformat(inv["due"]) + dt.timedelta(days=270) <= eom:
                self.post(eom, "WO " + inv["no"], f"Bad debt {inv['client']}", [("5900", inv["total"], 0), ("1100", 0, inv["total"])])
            else:
                still.append(inv)
        L["ar"] = still
        if self.game_rev:
            self.post(eom, f"GAMES-{self.tag}", "App store and Steam payouts (net of 30% fees)",
                      [("1000", self.game_rev, 0), ("4100", 0, self.game_rev)])
        if self.ad_rev:
            self.post(eom, f"ADS-{self.tag}", "Ad network revenue", [("1000", self.ad_rev, 0), ("4200", 0, self.ad_rev)])
        if self.pay_rev:
            self.post(eom, f"PAY-REV-{self.tag}", "CorvalisPay processing fees", [("1000", self.pay_rev, 0), ("4300", 0, self.pay_rev)])
            self.post(eom, f"PAY-COST-{self.tag}", "Interchange and card network fees",
                      [("5150", self.pay_cost, 0), ("1000", 0, self.pay_cost)])
        if self.fraud_loss:
            self.post(eom, f"CB-{self.tag}", "Lost chargebacks and fraud", [("5900", self.fraud_loss, 0), ("1000", 0, self.fraud_loss)])
        self.post(eom, f"PAYROLL-{self.tag}", "Payroll", [("5000", self.payroll, 0), ("1000", 0, self.payroll)])
        self.post(eom, f"PTAX-{self.tag}", "Employer payroll taxes and benefits", [("5050", self.taxes, 0), ("1000", 0, self.taxes)])
        p, g, m = self.w.load(self.ym)
        staff = len(self.staff)
        costs = {"5100": p * r.uniform(6000, 16000), "5200": 9000 + g * 38 + m * 22 + p * 900,
                 "5300": self.ua_spend + 15000, "5400": 32000 + staff * 520, "5500": staff * 240 + 6000,
                 "5600": 22000 + m * 30 + (48000 if self.ym[1] in (3, 9) else 0), "5700": 7000 + staff * 45,
                 "6000": staff * r.uniform(60, 160)}
        vendors = {"5100": "Northbridge Contractors", "5200": "Stratus Cloud", "5300": "Adreach Media",
                   "5400": "Lakeside Offices", "5500": "Toolstack Software", "5600": "Ledgerwise Advisory",
                   "5700": "Anchorpoint Insurance", "6000": "Talentway Recruiting"}
        bills = []
        for acct, amt in costs.items():
            amt = r2(amt)
            L["bill_no"] = L.get("bill_no", 0) + 1
            no = f"BILL-{L['bill_no']:06d}"
            bdate = self.d0 + dt.timedelta(days=r.randrange(self.days))
            bills.append([no, bdate.isoformat(), vendors[acct], ACCOUNTS[acct][0], amt])
            self.post(bdate, no, f"Bill {vendors[acct]}", [(acct, amt, 0), ("2000", 0, amt)])
            self.post(min(eom, bdate + dt.timedelta(days=30)), "PMT " + no, f"Payment {vendors[acct]}",
                      [("2000", amt, 0), ("1000", 0, amt)])
        if self.ym[1] in (1, 7):
            capex = r2(staff * 1900)
            self.post(self.d0, f"CAPEX-{self.tag}", "Laptops, devices and test hardware", [("1500", capex, 0), ("1000", 0, capex)])
        bal = L.setdefault("bal", {})
        dep = r2(bal.get("1500", 0) / 36)
        if dep:
            self.post(eom, f"DEP-{self.tag}", "Depreciation", [("5800", dep, 0), ("1510", 0, dep)])
        for _, _, a, _, _, dr, cr in self.gl:
            bal[a] = r2(bal.get(a, 0) + dr - cr)
        base = f"{F['fin']}/{self.y}"
        wb = new_book()
        add_sheet(wb, "Journal", ["Date", "Reference", "Account", "Account name", "Description", "Debit", "Credit"],
                  [[str(x[0])] + x[1:] for x in sorted(self.gl, key=lambda x: str(x[0]))], {"Description": 44, "Account name": 32})
        ws = wb["Journal"]
        n = ws.max_row
        ws[f"F{n + 1}"], ws[f"G{n + 1}"] = f"=SUM(F2:F{n})", f"=SUM(G2:G{n})"
        save_book(self.sink, f"{base}/General ledger/General ledger {self.tag}.xlsx", wb, "GL", BANNER)
        tb = [[a, ACCOUNTS[a][0], ACCOUNTS[a][1], max(bal.get(a, 0), 0), max(-bal.get(a, 0), 0)] for a in sorted(ACCOUNTS)]
        wb = new_book()
        add_sheet(wb, "Trial balance", ["Account", "Name", "Type", "Debit", "Credit"], tb, {"Name": 36})
        ws = wb["Trial balance"]
        n = ws.max_row
        ws[f"D{n + 1}"], ws[f"E{n + 1}"] = f"=SUM(D2:D{n})", f"=SUM(E2:E{n})"
        ws.cell(n + 2, 2, "Difference")
        ws[f"D{n + 2}"] = f"=D{n + 1}-E{n + 1}"
        save_book(self.sink, f"{base}/Trial balance/Trial balance {self.tag}.xlsx", wb, "TB", BANNER)
        wb = new_book()
        add_sheet(wb, "Client invoices", ["Invoice", "Date", "Client code", "Client", "Project", "Pricing", "Hours",
                                          "Amount USD", "Due"], invoices, {"Client": 30})
        add_sheet(wb, "Vendor bills", ["Bill", "Date", "Vendor", "Expense account", "Amount USD"], bills, {"Vendor": 26,
                                                                                                         "Expense account": 30})
        save_book(self.sink, f"{base}/Billing/Invoices and bills {self.tag}.xlsx", wb, "Billing", BANNER)
        aging = []
        for inv in L["ar"]:
            days = (eom - dt.date.fromisoformat(inv["due"])).days
            aging.append([inv["no"], inv["client"], inv["due"], inv["total"], max(days, 0),
                          "Current" if days <= 0 else "1-30" if days <= 30 else "31-60" if days <= 60 else "61-90" if days <= 90 else "90+"])
        wb = new_book()
        add_sheet(wb, "AR aging", ["Invoice", "Client", "Due", "Open USD", "Days overdue", "Bucket"], aging, {"Client": 30})
        save_book(self.sink, f"{base}/Receivables/AR aging {self.tag}.xlsx", wb, "AR aging", BANNER)
        pl = {a: r2(-sum(dr - cr for _, _, aa, _, _, dr, cr in self.gl if aa == a)) for a in ACCOUNTS
              if ACCOUNTS[a][1] in ("Revenue", "Expense")}
        L.setdefault("pl", {})[self.tag] = pl
        self.pl = pl
        if self.ym[1] == 12:
            self.year_end()

    def invoice_pdf(self, no, date, due, c, p, hours, amount):
        path = self.sink.path(f"{F['fin']}/{self.y}/Invoices/{self.tag}/{no} {safe(c['name'], 50)}.pdf")
        st = getSampleStyleSheet()
        doc = SimpleDocTemplate(path, pagesize=LETTER, title=f"Invoice {no}", author=CO, subject=BANNER)
        line = f"{p['name']} - {p['model']}" + (f" ({hours:,.1f} h at ${p['rate']:,.2f})" if p["model"] == "Time & materials" else "")
        t = Table([["Description", "Amount USD"], [line, f"{amount:,.2f}"], ["Total due", f"{amount:,.2f}"]], colWidths=[380, 100])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#243B55")),
                               ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, 1), 0.4, colors.grey),
                               ("ALIGN", (1, 0), (1, -1), "RIGHT"), ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold")]))
        doc.build([Paragraph(f"<b>{CO}</b>", st["Title"]), Paragraph(f"INVOICE {no}", st["Heading2"]),
                   Paragraph(f"Bill to: {c['name']}, {c['city']}, {c['country']}<br/>Attn: {c['contact']}<br/>"
                             f"Date: {date:%b %d, %Y} &nbsp; Due: {due:%b %d, %Y} &nbsp; Project: {p['code']}", st["Normal"]),
                   Spacer(1, 14), t, Spacer(1, 14), Paragraph(f"Pay by ACH quoting {no}. {BANNER}", st["Italic"])])
        self.sink.wrote(path)

    def year_end(self):
        L, y = self.L, self.ym[0]
        months = [f"{y}-{m:02d}" for m in range(1, 13) if f"{y}-{m:02d}" in L["pl"]]
        wb = new_book()
        ws = wb.create_sheet("Income statement")
        ws.append([f"{CO} - Income statement {y} (USD, fictional)"])
        ws.append(["Account", "Name"] + months + ["Full year"])
        last = chr(ord("C") + len(months) - 1)
        for kind, sign in (("Revenue", 1), ("Expense", -1)):
            for a in sorted(ACCOUNTS):
                if ACCOUNTS[a][1] == kind:
                    ws.append([a, ACCOUNTS[a][0]] + [sign * L["pl"][m].get(a, 0) for m in months])
                    ws[f"{chr(ord(last) + 1)}{ws.max_row}"] = f"=SUM(C{ws.max_row}:{last}{ws.max_row})"
        ws.column_dimensions["B"].width = 36
        bs = wb.create_sheet("Balance sheet")
        bs.append([f"{CO} - Balance sheet at Dec 31 {y} (USD, fictional; before year-end close)"])
        for a in sorted(ACCOUNTS):
            if ACCOUNTS[a][1] in ("Asset", "Liability", "Equity"):
                bs.append([a, ACCOUNTS[a][0], ACCOUNTS[a][1], L["bal"].get(a, 0) * (1 if ACCOUNTS[a][1] == "Asset" else -1)])
        bs.column_dimensions["B"].width = 36
        save_book(self.sink, f"{F['fin']}/{y}/Financial statements {y}.xlsx", wb, "Financial statements", BANNER)
        prev = L.get("pl", {})
        wb = new_book()
        ws = wb.create_sheet("Budget")
        ws.append([f"{CO} - Operating budget {y + 1} (USD, fictional)"])
        ws.append(["Growth assumption", 0.15])
        ws.append(["Account", "Name", f"{y} actual"] + [dt.date(y + 1, m, 1).strftime("%b") for m in range(1, 13)] + ["Total"])
        for a in sorted(ACCOUNTS):
            if ACCOUNTS[a][1] in ("Revenue", "Expense"):
                act = r2(sum(abs(prev[m].get(a, 0)) for m in months))
                row = ws.max_row + 1
                ws.append([a, ACCOUNTS[a][0], act] + [f"=$C{row}*(1+$B$2)/12" for _ in range(12)] + [f"=SUM(D{row}:O{row})"])
        ws.column_dimensions["B"].width = 34
        save_book(self.sink, f"{F['fin']}/{y + 1}/Budget {y + 1}.xlsx", wb, "Budget", BANNER)

    # ---------------- reports, meetings, mail
    def reports(self):
        r = self.rng
        p, g, m = self.w.load(self.ym)
        rev = sum(v for a, v in self.pl.items() if ACCOUNTS[a][1] == "Revenue")
        exp = -sum(v for a, v in self.pl.items() if ACCOUNTS[a][1] == "Expense")
        wb = new_book()
        add_sheet(wb, "Company KPIs", ["Measure", "Value"], [
            ["Month", self.tag], ["Active client projects", p], ["Daily players (avg, all games)", int(g * 1000)],
            ["Active merchants", m], ["Headcount", len(self.staff)], ["Revenue USD", r2(rev)], ["Expenses USD", r2(exp)],
            ["Net USD", r2(rev - exp)], ["Production incidents", len(self.incidents)], ["Support tickets", self.tickets_n],
            ["Cash at bank USD", self.L["bal"].get("1000", 0)]], {"Measure": 34})
        save_book(self.sink, f"{F['rep']}/{self.y}/KPI dashboard {self.tag}.xlsx", wb, "KPIs", BANNER)
        if self.ym[1] % 3 == 0:
            q = self.ym[1] // 3
            d = new_doc(f"Board update Q{q} {self.y}", f"{CO} | quarter ending {self.d0:%B %Y}")
            ms = [ym_s(ym_add(self.ym, -k)) for k in (2, 1, 0)]
            rows = []
            for mm in ms:
                pl = self.L["pl"].get(mm, {})
                a = sum(v for k, v in pl.items() if ACCOUNTS[k][1] == "Revenue")
                b = -sum(v for k, v in pl.items() if ACCOUNTS[k][1] == "Expense")
                rows.append([mm, r2(a), r2(b), r2(a - b)])
            doc_table(d, ["Month", "Revenue", "Expenses", "Net"], rows)
            d.add_heading("Highlights", 1)
            for x in r.sample(["Two new game publisher clients signed", "CorvalisPay passed its annual PCI DSS assessment",
                               "Puzzle Harbor reached a record weekend DAU", "Hiring slowed to plan",
                               "Chargeback rate fell after new risk rules", "Ironvale Arena season 3 launched",
                               "Bank client renewed for two years"], 3):
                bullet(d, x)
            d.add_heading("Risks", 1)
            for x in r.sample(["App store policy changes on payments", "Client concentration in top five accounts",
                               "Rising user acquisition costs", "Card network fee increases", "Senior engineer attrition"], 2):
                bullet(d, x)
            save_doc(self.sink, f"{F['rep']}/{self.y}/Board updates/Board update {self.y} Q{q}.docx", d)

    def meetings(self):
        r = self.rng
        heads = [e for e in self.staff if e["role"] in HEADS.values()]
        if not heads:
            return
        base = f"{F['mail']}/{self.y}"
        for day in range(0, self.days, 7):
            date = self.d0 + dt.timedelta(days=day)
            lines = [(heads[0]["name"], "Morning. Quick round: games, fintech, delivery, people."),
                     (next((e["name"] for e in heads if e["role"] == "Head of Games"), heads[0]["name"]),
                      r.choice(["Weekend DAU was strong, the event drove it.", "Crash rate is back under one percent.",
                                "We are testing a cheaper starter pack this week."])),
                     (next((e["name"] for e in heads if e["role"] == "Head of Fintech"), heads[0]["name"]),
                      r.choice(["Two merchants in review for high dispute rates.", "Payout times are within target.",
                                "The new risk rules cut card testing by half."])),
                     (next((e["name"] for e in heads if e["role"] == "Chief Technology Officer"), heads[0]["name"]),
                      f"{len(self.projects)} client projects active; {sum(1 for i in self.incidents if i[1] in ('SEV1', 'SEV2'))} serious incidents this month."),
                     (next((e["name"] for e in heads if e["role"] == "Head of People"), heads[0]["name"]),
                      r.choice(["Three offers out for engineers.", "Review cycle starts next week.", "Attrition is on plan."])),
                     (heads[0]["name"], "Thanks. Actions in the doc.")]
            S.zoom_export(self, f"{base}/Zoom Meetings/{date} 0930 Leadership weekly", "Leadership weekly",
                          dt.datetime.combine(date, dt.time(9, 30)), [(e["name"], e["email"], False) for e in heads], lines,
                          heads[0]["name"], r.randint(30, 50))
        for p in self.projects[:12]:
            team = self.team.get(p["id"], [])[:5]
            if not team:
                continue
            c = self.cl_by[p["client"]]
            start = dt.datetime.combine(self.d0 + dt.timedelta(days=min(self.days - 1, 13 + p["id"] % 10)), dt.time(16, 0))
            lines = [(team[0]["name"], f"Sprint review for {p['name']}. We closed most of the sprint goals."),
                     (c["contact"], r.choice(["The new build feels much faster.", "Can we see the payment flow on staging?",
                                              "We need the release before the holiday sale."])),
                     (team[-1]["name"], r.choice(["Staging is updated, link in the chat.", "We'll prioritise that next sprint.",
                                                  "There is one blocker on the client API keys."])),
                     (team[0]["name"], "Retro notes are in the shared doc. Thanks everyone.")]
            S.meet_export(self, f"{base}/Google Meet/{start:%Y-%m-%d} Sprint review {safe(p['code'], 20)}",
                          f"Sprint review {p['code']}", start,
                          [(e["name"], e["email"], False) for e in team] + [(c["contact"], f"pm@{c['domain']}", True)],
                          lines, r.randint(30, 60))
        for p in self.projects[:40]:
            c = self.cl_by[p["client"]]
            self.mail(self.ts(), f"delivery@{DOMAIN}", f"pm@{c['domain']}", f"[{p['code']}] Weekly update",
                      f"Hi {c['contact'].split()[0]},\n\nThis week on {p['name']}: sprint goals met, next demo on Thursday.\n\n"
                      f"Thanks,\n{CO} Delivery\n")

    def write_mail(self):
        for row in getattr(self, "fraud_alerts", [])[:30]:
            self.mail(self.ts(), f"risk@{DOMAIN}", f"payments-ops@{DOMAIN}", f"Fraud alert {row[1]} ({row[5]})",
                      f"Rule {row[5]} fired for merchant {row[2]}: amount ${row[3]:,.2f}, risk {row[4]}. Outcome: {row[6]}.\n")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for i, (t, subj, raw) in enumerate(sorted(self.emails, key=lambda x: x[0]), 1):
                z.writestr(f"{i:05d} {safe(subj, 70)}.eml", raw)
        self.sink.write_bytes(f"{F['mail']}/{self.y}/Email exports/Mailbox export {self.tag}.zip", buf.getvalue())


# ---------------------------------------------------------------- static documents & snapshots

POLICIES = {
    "Information Security Policy": ["Scope", "Roles", "Access control", "Encryption", "Secure development", "Logging",
                                    "Vendor security", "Incident response", "Exceptions"],
    "PCI DSS Compliance Program": ["Cardholder data environment", "Network segmentation", "Key management",
                                   "Vulnerability management", "Quarterly scans and annual assessment", "Evidence"],
    "AML and KYC Program": ["Customer due diligence", "Beneficial ownership", "Sanctions screening",
                            "Transaction monitoring", "Suspicious activity escalation", "Record keeping"],
    "Chargeback and Dispute Handling": ["Notification", "Evidence collection", "Representment", "Merchant reserves",
                                        "Reporting"],
    "Privacy Policy (Games)": ["Data we collect", "How we use it", "Children's privacy (COPPA)", "Your rights",
                               "Retention", "Contact"],
    "Terms of Service (Games)": ["Accounts", "Virtual items and currency", "Purchases and refunds", "Conduct",
                                 "Termination", "Disputes"],
    "Data Retention Policy": ["Player data", "Payment records", "Employee records", "Source code", "Deletion"],
    "Code of Conduct": ["Respect", "Conflicts of interest", "Gifts", "Speaking up", "Discipline"],
    "Remote Work Policy": ["Eligibility", "Equipment", "Security at home", "Working hours", "Expenses"],
    "Expense Policy": ["What we reimburse", "Approval", "Receipts", "Travel", "Corporate cards"],
    "Release Management Policy": ["Branching", "Code review", "Testing gates", "Change approval", "Rollback"],
    "Responsible Gaming and Monetisation Policy": ["Fair odds disclosure", "Spending limits", "Minors", "Ad standards"],
}
POLICY_TEXT = ["{co} applies this to all employees, contractors and systems in scope.",
               "Owners review this every year and after any material change.",
               "Evidence is kept for at least three years and made available to auditors.",
               "Exceptions need written approval from the policy owner and expire after twelve months.",
               "Training on this topic is part of onboarding and is refreshed yearly.",
               "Breaches are reported to the policy owner within one business day."]
IT_SOPS = {
    "SOP-ENG-01 Code review and merge": [("Steps", ["Open a pull request linked to its issue.", "At least one reviewer; two for payments and auth code.",
                                                    "CI must pass: unit, integration, security scan.", "Squash-merge to main; delete the branch."])],
    "SOP-ENG-02 Production release": [("Steps", ["Cut a release branch on Tuesday.", "Run the regression suite and load test.",
                                                  "Deploy to canary (5%) for one hour.", "Watch error rate and latency SLOs.",
                                                  "Roll forward to 100% or roll back in one command."])],
    "SOP-ENG-03 On-call and incident command": [("Steps", ["Page acknowledges within 5 minutes.", "SEV1/SEV2: open an incident channel and a Meet bridge.",
                                                           "Assign incident commander, communicator and scribe.",
                                                           "Update the status page every 30 minutes.", "Postmortem within 5 business days."])],
    "SOP-PAY-01 Merchant onboarding (KYB)": [("Steps", ["Verify business registration and beneficial owners.", "Screen sanctions and PEP lists.",
                                                         "Review website, refund policy and products.", "Assign risk tier and reserve.",
                                                         "Compliance officer approves high-risk merchants."])],
    "SOP-PAY-02 Daily settlement and reconciliation": [("Steps", ["Import processor settlement files at 06:00.",
                                                                   "Match captured transactions to settlement lines.",
                                                                   "Investigate breaks over $50 the same day.",
                                                                   "Release merchant payouts after reconciliation is signed off."])],
    "SOP-PAY-03 Chargeback response": [("Steps", ["Notify the merchant within 24 hours.", "Collect evidence: receipt, delivery, logs.",
                                                   "Submit representment before the network deadline.", "Record outcome and update risk rules."])],
    "SOP-GAME-01 Live-ops event": [("Steps", ["Plan event and rewards two sprints ahead.", "Economy review by design and data.",
                                               "Feature-flag the event; test on staging.", "Monitor KPIs hourly on day one."])],
    "SOP-SEC-01 Access review": [("Steps", ["Quarterly export of access to production, payments and admin tools.",
                                             "Managers confirm each grant.", "Remove unconfirmed access within 5 days."])],
    "SOP-PRIV-01 Player data request": [("Steps", ["Verify the requester owns the account.", "Export or delete within 30 days.",
                                                    "Keep payment records required by law.", "Log the request and outcome."])],
    "SOP-HR-01 Employee onboarding": [("Steps", ["Send equipment one week before start.", "Create accounts from the role template.",
                                                  "Assign a buddy and 30/60/90-day goals.", "Complete security and privacy training in week one."])],
}


def static_docs(w, sink):
    r = random.Random(SEED + 5)
    for i, (title, secs) in enumerate(POLICIES.items(), 1):
        d = new_doc(title, f"{CO} | Document CL-POL-{i:02d} | Version {r.randint(1, 5)}.{r.randint(0, 9)}")
        for j, s in enumerate(secs, 1):
            d.add_heading(f"{j}. {s}", 1)
            for t in r.sample(POLICY_TEXT, 3):
                d.add_paragraph(t.format(co=CO))
        save_doc(sink, f"{F['pol']}/{title}.docx", d)
    S.it_sops(w, sink, sys.modules[__name__])
    for role, (dept, base, _) in ROLES.items():
        d = new_doc(f"Job description - {role}", f"Department: {dept}")
        doc_table(d, ["Reports to", "Salary band (USD/year)"], [[HEADS.get(dept, "Chief Executive Officer"),
                                                                 f"{base * 12 * .88:,.0f} - {base * 12 * 1.2:,.0f}"]])
        d.add_heading("What you will do", 1)
        for x in r.sample(["Ship high-quality work in two-week sprints", "Work closely with clients and product",
                           "Keep player and payment data safe", "Write clear documentation", "Mentor teammates",
                           "Measure outcomes, not output", "Take part in on-call or live-ops rotations"], 4):
            bullet(d, x)
        save_doc(sink, f"{F['hr']}/Job descriptions/{role}.docx", d)
    d = new_doc(f"{CO} - dataset README", BANNER)
    d.add_paragraph("Synthetic, internally consistent records of a fictional game and fintech studio: client projects with "
                    "contracts, statements of work and invoices; engineering issues, commits, builds and incidents; live-game "
                    "KPIs, purchases and telemetry; payment transactions, disputes and merchant onboarding; support; marketing; "
                    "payroll and per-person HR files; double-entry books whose trial balance balances every month.")
    for v in F.values():
        bullet(d, v.split("/", 1)[1])
    save_doc(sink, f"{ROOT}/README.docx", d)
    wb = new_book()
    add_sheet(wb, "Chart of accounts", ["Account", "Name", "Type"], [[a, *v] for a, v in sorted(ACCOUNTS.items())], {"Name": 36})
    add_sheet(wb, "Games", ["Game", "Genre", "Platform", "Launch", "Peak DAU (k)", "ARPDAU USD"],
              [[g[0], g[1], g[2], ym_s(g[3]), g[4], g[5]] for g in GAMES], {"Game": 22})
    add_sheet(wb, "Notice", ["Notice"], [[BANNER], ["Domains use the reserved .example TLD; card data is never real."]], {"Notice": 100})
    save_book(sink, f"{ROOT}/Data dictionary.xlsx", wb, "Data dictionary", BANNER)


def snapshots(w, ym, sink):
    rows = [[c["code"], c["name"], c["kind"], c["city"], c["country"], ym_s(c["since"]),
             "Lost" if c["end"] and c["end"] <= ym else "Active", c["contact"], f"pm@{c['domain']}"]
            for c in w.clients if c["since"] <= ym]
    projects = [[p["code"], p["name"], w.clients[p["client"]]["name"], p["type"], p["model"], ym_s(p["start"]), ym_s(p["end"]),
                 p["team"], p["rate"], p["budget"], p["risk"], "Complete" if p["end"] < ym else "Active" if p["start"] <= ym else "Signed"]
                for p in w.projects if p["start"] <= ym_add(ym, 1)]
    wb = new_book()
    add_sheet(wb, "Clients", ["Code", "Client", "Segment", "City", "Country", "Since", "Status", "Contact", "Email"], rows,
              {"Client": 30, "Contact": 22, "Email": 28})
    add_sheet(wb, "Projects", ["Project", "Name", "Client", "Type", "Pricing", "Start", "End", "Team", "Rate", "Budget USD",
                               "Risk", "Status"], projects, {"Name": 34, "Client": 30, "Type": 28})
    save_book(sink, f"{F['crm']}/Client and project register (as of {ym_s(ym)}).xlsx", wb, "Register", BANNER)
    merch = [[m["code"], m["name"], m["category"], m["mcc"], m["city"], m["country"], ym_s(m["since"]), m["risk"],
              f"{m['rate'] * 100:.1f}% + ${m['fixed']:.2f}", "Closed" if m["end"] and m["end"] <= ym else "Active"]
             for m in w.merchants if m["since"] <= ym]
    wb = new_book()
    add_sheet(wb, "Merchants", ["Merchant", "Name", "Category", "MCC", "City", "Country", "Since", "Risk tier", "Pricing",
                                "Status"], merch, {"Name": 28, "Category": 22})
    save_book(sink, f"{F['pay']}/Merchant register (as of {ym_s(ym)}).xlsx", wb, "Merchants", BANNER)
    staff = [[e["id"], e["name"], e["email"], e["role"], e["level"], e["dept"], e["manager"], e["nationality"], ym_s(e["hired"]),
              ym_s(e["left"]) if e["left"] and e["left"] <= ym else "", w.pay(e, ym)] for e in w.staff if e["hired"] <= ym]
    wb = new_book()
    add_sheet(wb, "Employees", ["Employee", "Name", "Email", "Role", "Level", "Department", "Manager", "Nationality", "Hired",
                                "Left", "Base USD/month"], staff, {"Name": 22, "Email": 34, "Role": 26})
    save_book(sink, f"{F['hr']}/Employee roster (as of {ym_s(ym)}).xlsx", wb, "Roster", BANNER)


def build(sink, st, max_months=24, scale=1.0, detail=1.0):
    today = dt.date.today()
    last_full = ym_add((today.year, today.month), -1)
    done = tuple(st["done"]) if st.get("done") else None
    nxt = ym_add(done, 1) if done else START
    if nxt > last_full:
        return 0
    end = min(last_full, ym_add(nxt, max_months - 1))
    w = World(last_full, scale)
    ledger = st.setdefault("ledger", {})
    if not st.get("static"):
        static_docs(w, sink)
        st["static"] = True
    n, ym = 0, nxt
    while ym <= end and not sink.full():
        Month(w, ym, ledger, sink, detail).run()
        st["done"] = list(ym)
        n += 1
        ym = ym_add(ym, 1)
    if n:
        last = tuple(st["done"])
        snapshots(w, last, sink)
        if last[1] != 12:
            S.person_year(w, ledger, last[0], sink, sys.modules[__name__], final=False)
    return n
