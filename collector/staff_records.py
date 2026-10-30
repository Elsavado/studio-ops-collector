"""Per-person HR and accounting files, Zoom / Google Meet meeting exports and IT SOPs for the fictional company.

Everything here is synthetic. Per-person monthly figures are kept in the ledger state during a year and written
into each employee's file at year end (and to date at the end of a run)."""
import datetime as dt
import io
import random

from .common import add_sheet, new_book, safe, save_book

EXPENSES = {  # role: [(category, description, low, high, monthly chance)]
    "Field Technician": [("Mileage", "Site visits by private car", 40, 260, .55), ("Parking & tolls", "Client sites", 8, 60, .5),
                         ("Tools & consumables", "Cable ties, connectors, cleaning kit", 15, 120, .3)],
    "Account Manager": [("Client entertainment", "Lunch with client", 60, 420, .5), ("Travel", "Flights and hotel for client visit", 300, 2400, .25),
                        ("Taxi", "Client meetings", 20, 140, .6)],
    "Project Manager": [("Travel", "Installation site visit", 200, 1800, .35), ("Taxi", "Site meetings", 15, 120, .5)],
    "Network Engineer": [("Certification", "Exam voucher", 150, 450, .06), ("Taxi", "Out-of-hours call-out", 20, 90, .3)],
    "Cloud Engineer": [("Certification", "Cloud certification exam", 150, 400, .06), ("Software", "Personal licence for lab", 10, 60, .15)],
    "SOC Analyst": [("Meals", "Night-shift meal allowance claim", 10, 45, .2), ("Taxi", "Late-shift transport", 15, 60, .25)],
    "SOC Shift Lead": [("Meals", "Shift meal", 10, 45, .2), ("Taxi", "Late-shift transport", 15, 60, .25)],
    "IT Support Specialist": [("Taxi", "Client site support", 15, 80, .25)],
}
MGMT_EXP = [("Travel", "Client and partner visit", 400, 4200, .35), ("Client entertainment", "Dinner with client", 120, 900, .3)]

TRAINING = {
    "SOC Analyst": ["CCTV operator certification", "Video analytics tuning", "Incident escalation drill", "CompTIA Security+"],
    "SOC Shift Lead": ["Major incident management", "Leadership on shift", "CompTIA CySA+"],
    "Field Technician": ["Working at height", "IP camera installation", "Electrical safety", "Marine site safety induction"],
    "Network Engineer": ["Cisco CCNA", "Firewall administration", "Network segmentation for IoT"],
    "Cloud Engineer": ["Cloud practitioner", "Kubernetes administration", "Backup and recovery"],
    "IT Support Specialist": ["ITIL 4 Foundation", "Customer service for help desks", "VMS administration"],
    "Account Manager": ["Consultative selling", "Contract basics"], "Project Manager": ["PRINCE2 Foundation", "Site safety management"],
}
COMMON_TRAINING = ["Data protection awareness", "Phishing awareness", "Fire safety and first aid"]


def collect(m):
    """After the month's generators ran: per-person work counts, pay and expenses into the ledger."""
    L = m.L
    eid = {e["name"]: e["id"] for e in m.staff}
    work = {}
    for a in m.alerts:
        k = a[6]["id"]
        work.setdefault(k, [0, 0, 0])[0] += 1
    for t in m.tickets:
        if t[6] in eid:
            work.setdefault(eid[t[6]], [0, 0, 0])[1] += 1
    for w in m.workorders:
        if w[7] in eid:
            work.setdefault(eid[w[7]], [0, 0, 0])[2] += 1
    people = L.setdefault("people", {})
    for e in m.staff:
        base, house, trans, ot, otp, gross, eos, leave = m.pay_lines[e["id"]]
        al, tk, wo = work.get(e["id"], [0, 0, 0])
        rec = people.setdefault(e["id"], {})
        rec[m.tag] = [base, house, trans, ot, otp, gross, eos, leave, al, tk, wo, m.expense_by.get(e["id"], 0.0)]


def expenses(m):
    """Expense claims for the month: register in Finance, posted to the GL as reimbursed through the bank."""
    r = random.Random(f"exp-{m.tag}")
    rows, m.expense_by = [], {}
    mgr = {e["id"]: e for e in m.staff}
    heads = {"Chief Executive Officer", "Operations Director", "Finance Manager", "SOC Manager", "Field Services Manager"}
    for e in m.staff:
        opts = EXPENSES.get(e["role"]) or (MGMT_EXP if e["role"] in heads else [])
        for cat, desc, lo, hi, p in opts:
            if r.random() < p:
                amt = round(r.uniform(lo, hi), 2)
                day = r.randrange(m.days)
                appr = mgr.get(e["manager"], {}).get("name", "Finance Manager")
                rows.append([f"EXP-{m.tag.replace('-', '')}-{len(rows) + 1:05d}", e["id"], e["name"], e["dept"],
                             (m.d0 + dt.timedelta(days=day)).isoformat(), cat, desc, amt,
                             "Yes" if r.random() < .93 else "Missing", appr,
                             (m.d0 + dt.timedelta(days=min(m.days - 1, day + r.randint(3, 12)))).isoformat()])
                m.expense_by[e["id"]] = round(m.expense_by.get(e["id"], 0) + amt, 2)
    total = round(sum(x[7] for x in rows), 2)
    if total:
        eom = m.d0 + dt.timedelta(days=m.days - 1)
        m.post(eom, f"EXP-{m.tag}", "Staff expense claims reimbursed", [("6200", total, 0), ("1000", 0, total)])
    wb = new_book()
    add_sheet(wb, "Expense claims", ["Claim", "Employee", "Name", "Department", "Date", "Category", "Description",
                                     "Amount USD", "Receipt", "Approved by", "Paid"], rows,
              {"Name": 22, "Description": 40, "Approved by": 22, "Department": 20})
    ws = wb["Expense claims"]
    n = ws.max_row
    ws.cell(n + 1, 7, "TOTAL")
    ws[f"H{n + 1}"] = f"=SUM(H2:H{n})"
    save_book(m.sink, f"{m.F['fin']}/{m.y}/Expense claims/Expense claims {m.tag}.xlsx", wb, "Expense claims", m.BANNER)
    m.expense_rows = rows


# ---------------------------------------------------------------- per-person files

def person_dir(m_or_F, e):
    return f"{m_or_F['hr']}/Employee Files/{e['id']} {safe(e['name'], 40)}"


def joiner_docs(m, e):
    C, d0 = m.C, m.d0
    base = person_dir(m.F, e)
    pay = m.w.pay(e, m.ym)
    mgr = next((x for x in m.w.staff if x["id"] == e["manager"]), None)
    d = C.new_doc(f"Offer of employment - {e['name']}", f"{C.CO} | {d0 - dt.timedelta(days=21):%d %B %Y}")
    d.add_paragraph(f"Dear {e['name'].split()[0]},")
    d.add_paragraph(f"We are pleased to offer you the position of {e['role']} in our {e['dept']} team, starting on "
                    f"{d0:%d %B %Y}. You will report to {mgr['name'] if mgr else 'the Chief Executive Officer'}.")
    C.doc_table(d, ["Item", "Terms"], [["Basic salary", f"USD {pay:,.2f} per month"],
                                       ["Housing allowance", f"USD {pay * .25:,.2f} per month"],
                                       ["Transport allowance", f"USD {pay * .08:,.2f} per month"],
                                       ["Probation", "6 months"], ["Annual leave", "30 calendar days"],
                                       ["Notice period", "1 month (3 months for managers)"],
                                       ["End-of-service benefit", "21 days' basic salary per year of service"]])
    d.add_paragraph("Please sign and return a copy of this letter to accept the offer.")
    C.save_doc(m.sink, f"{base}/01 Contract & Onboarding/Offer letter.docx", d)
    d = C.new_doc(f"Employment contract - {e['name']}", f"Employee {e['id']}")
    for h, t in [("1. Position", f"The Employee is employed as {e['role']} ({e['dept']})."),
                 ("2. Start date", f"Employment starts on {d0:%d %B %Y}."),
                 ("3. Working hours", "48 hours per week; SOC roles work rotating 8-hour shifts on a four-crew pattern."),
                 ("4. Pay", f"Basic USD {pay:,.2f}, housing 25% and transport 8% of basic, paid monthly in arrears."),
                 ("5. Confidentiality", "Client video, personal data and security details are confidential during "
                                        "and after employment."),
                 ("6. IT and security", "The Employee follows the Information Security, Acceptable Use and Remote "
                                        "Access policies and completes annual security training."),
                 ("7. Termination", "Either party may end the contract with the notice in the offer letter.")]:
        d.add_heading(h, 1)
        d.add_paragraph(t)
    C.doc_table(d, ["Signed for the company", "Employee"], [[m.w.staff[0]["name"], e["name"]]])
    C.save_doc(m.sink, f"{base}/01 Contract & Onboarding/Employment contract.docx", d)
    d = C.new_doc(f"IT onboarding checklist - {e['name']}", f"Ticket HR-ONB-{e['id']}")
    items = ["Directory account created and added to department group", f"Mailbox {e['email']} created",
             "MFA enrolled on company phone", "Laptop issued and encrypted (asset tag recorded)",
             "Password manager vault shared", "Ticketing system account", "Security awareness module assigned"]
    if "SOC" in e["role"]:
        items += ["VMS operator account (role: monitoring)", "Video wall workstation login", "Alarm platform login"]
    if e["role"] in ("Field Technician", "Network Engineer"):
        items += ["Recorder admin credentials via vault (time-limited)", "VPN profile", "Service van fuel card"]
    if e["role"] == "Cloud Engineer":
        items += ["Cloud console access via SSO with just-in-time elevation", "Infrastructure repo access"]
    C.doc_table(d, ["Step", "Done by", "Date"], [[x, "IT Support", (d0 + dt.timedelta(days=i % 3)).isoformat()]
                                                for i, x in enumerate(items)])
    C.save_doc(m.sink, f"{base}/01 Contract & Onboarding/IT onboarding checklist.docx", d)


def leaver_docs(m, e):
    C = m.C
    base = person_dir(m.F, e)
    pay = m.w.pay(e, m.ym)
    months = (m.ym[0] - e["hired"][0]) * 12 + m.ym[1] - e["hired"][1]
    years = round(months / 12, 2)
    gratuity = round(pay * 21 / 30 * years, 2) if months >= 12 else 0.0
    last = m.d0 - dt.timedelta(days=1)
    d = C.new_doc(f"Confirmation of end of employment - {e['name']}", f"{C.CO} | {last:%d %B %Y}")
    d.add_paragraph(f"This confirms that your employment as {e['role']} ends on {last:%d %B %Y} "
                    f"({e['leave_reason'].lower()}). Your final settlement is attached.")
    C.save_doc(m.sink, f"{base}/07 Exit/Exit letter.docx", d)
    wb = new_book()
    ws = add_sheet(wb, "Final settlement", ["Item", "Value"], [
        ["Employee", f"{e['id']} {e['name']}"], ["Role", e["role"]], ["Hired", C.ym_s(e["hired"])],
        ["Last day", last.isoformat()], ["Service (years)", years], ["Basic salary USD", pay],
        ["End-of-service days per year", 21], ["Gratuity USD", gratuity],
        ["Leave encashment USD", round(pay / 30 * random.Random(e["id"]).randint(0, 12), 2)]])
    ws["A11"], ws["B11"] = "Total payable USD", "=B9+B10"
    save_book(m.sink, f"{base}/07 Exit/Final settlement.xlsx", wb, "Final settlement", m.BANNER)
    d = C.new_doc(f"IT offboarding checklist - {e['name']}")
    C.doc_table(d, ["Step", "Status"], [[x, "Done"] for x in [
        "Directory account disabled at 17:00 on last day", "Mailbox converted to shared, forwarded to manager",
        "MFA tokens revoked", "VPN and VMS accounts removed", "Shared vault entries rotated",
        "Laptop and phone returned and wiped", "Access cards deactivated on all client sites"]])
    C.save_doc(m.sink, f"{base}/07 Exit/IT offboarding checklist.docx", d)


def person_year(w, L, y, sink, C, final=True):
    """Write each employee's pay, expenses, work log, training and review for year y."""
    people = L.get("people", {})
    by_id = {e["id"]: e for e in w.staff}
    F, BANNER = C.F, C.BANNER
    for eid, recs in people.items():
        months = sorted(k for k in recs if k.startswith(f"{y}-"))
        if not months or eid not in by_id:
            continue
        e = by_id[eid]
        base = person_dir(F, e)
        rows = [[k] + recs[k][:8] for k in months]
        wb = new_book()
        ws = add_sheet(wb, f"Pay {y}", ["Month", "Basic", "Housing", "Transport", "OT hours", "OT pay", "Gross",
                                        "EOS accrual", "Leave days"], rows)
        n = ws.max_row
        ws.cell(n + 1, 1, "TOTAL")
        for col in "BCDEFGHI":
            ws[f"{col}{n + 1}"] = f"=SUM({col}2:{col}{n})"
        save_book(sink, f"{base}/02 Pay/Pay statement {y}.xlsx", wb, "Pay statement", BANNER)
        wl = [[k, recs[k][8], recs[k][9], recs[k][10], recs[k][3], recs[k][7]] for k in months]
        wb = new_book()
        add_sheet(wb, f"Work {y}", ["Month", "SOC alarms handled", "Help desk tickets", "Work orders",
                                    "Overtime hours", "Leave days"], wl)
        save_book(sink, f"{base}/06 Work Log/Work log {y}.xlsx", wb, "Work log", BANNER)
        if any(recs[k][11] for k in months):
            wb = new_book()
            add_sheet(wb, f"Expenses {y}", ["Month", "Claims total USD"], [[k, recs[k][11]] for k in months])
            save_book(sink, f"{base}/05 Expenses/Expense claims {y}.xlsx", wb, "Expenses", BANNER)
        r = random.Random(f"{eid}-{y}")
        courses = r.sample(TRAINING.get(e["role"], []) + COMMON_TRAINING, min(3, len(TRAINING.get(e["role"], [])) + 3))
        tr = []
        for c in courses:
            when = dt.date(y, r.randint(int(months[0][5:]), int(months[-1][5:])), r.randint(1, 28))
            tr.append([c, "Brightpath Training" if "awareness" not in c.lower() else "Online module", when.isoformat(),
                       r.choice([4, 8, 16, 24, 40]), r.choices(["Passed", "Passed", "Passed", "Failed - resit"])[0],
                       (when + dt.timedelta(days=365 * r.choice([1, 2, 3]))).isoformat()])
        wb = new_book()
        add_sheet(wb, f"Training {y}", ["Course", "Provider", "Date", "Hours", "Result", "Valid until"], tr,
                  {"Course": 34, "Provider": 22})
        save_book(sink, f"{base}/04 Training/Training record {y}.xlsx", wb, "Training", BANNER)
        if final and len(months) >= 6:
            peers = [sum(v[8] + v[9] + v[10] for kk, v in rr.items() if kk.startswith(f"{y}-"))
                     for pid, rr in people.items() if pid in by_id and by_id[pid]["role"] == e["role"]]
            mine = sum(recs[k][8] + recs[k][9] + recs[k][10] for k in months)
            avg = sum(peers) / max(len(peers), 1)
            score = 3 if not avg else (5 if mine > avg * 1.25 else 4 if mine > avg * 1.05 else 3 if mine > avg * .8 else 2)
            score = max(1, min(5, score + r.choice([-1, 0, 0, 0, 1])))
            mgr = by_id.get(e["manager"], {"name": "Chief Executive Officer"})
            d = C.new_doc(f"Performance review {y} - {e['name']}", f"{e['role']} | Reviewer: {mgr['name']}")
            C.doc_table(d, ["Measure", "Value"], [["Months in year", len(months)],
                                                  ["Alarms / tickets / work orders", mine],
                                                  ["Team average", round(avg)], ["Overtime hours",
                                                                                sum(recs[k][3] for k in months)],
                                                  ["Overall rating (1-5)", score]])
            d.add_heading("Strengths", 1)
            for x in r.sample(["Calm and accurate under pressure", "Clear written handovers", "Strong client rapport",
                               "Helps new colleagues", "Good root-cause thinking", "Reliable attendance"], 2):
                C.bullet(d, x)
            d.add_heading("Development goals", 1)
            for x in r.sample(courses + ["Reduce time to acknowledge critical alarms", "Lead one process improvement",
                                         "Shadow the shift lead role"], 2):
                C.bullet(d, x)
            d.add_paragraph({5: "Exceptional year; recommended for promotion review.",
                             4: "Exceeded expectations.", 3: "Met expectations.",
                             2: "Partly met expectations; improvement plan agreed.",
                             1: "Did not meet expectations; formal plan in place."}[score])
            C.save_doc(sink, f"{base}/03 Performance/Performance review {y}.docx", d)
    if final:
        for recs in people.values():
            for k in [k for k in recs if k.startswith(f"{y}-")]:
                del recs[k]
        for k in [k for k, v in people.items() if not v]:
            del people[k]


# ---------------------------------------------------------------- Zoom / Google Meet exports

def _vtt_time(s):
    return f"{int(s // 3600):02d}:{int(s % 3600 // 60):02d}:{s % 60:06.3f}"


def zoom_export(m, folder, title, start, people, lines, host, minutes=40):
    """Zoom cloud-recording exports: transcript .vtt, chat .txt and participants .csv."""
    r = m.rng
    stamp = f"GMT{start:%Y%m%d-%H%M%S}"
    mid = f"{r.randrange(10 ** 10, 10 ** 11)}"
    t, vtt, chat = 2.0, ["WEBVTT", ""], []
    gap = minutes * 60 / max(len(lines), 1)
    for i, (who, text) in enumerate(lines, 1):
        dur = max(2.5, len(text) / 14)
        vtt += [str(i), f"{_vtt_time(t)} --> {_vtt_time(t + dur)}", f"{who}: {text}", ""]
        t += max(dur + 0.5, gap * r.uniform(0.6, 1.4))
        if r.random() < 0.08:
            chat.append(f"{(start + dt.timedelta(seconds=t)):%H:%M:%S} From {who} to Everyone:\n\t"
                        + r.choice(["Link to the dashboard is in the channel", "+1", "I'll take that action",
                                    "Can you share the slide again?", "Sharing the ticket number now"]))
    end = start + dt.timedelta(seconds=t + 30)
    m.sink.write_text(f"{folder}/{stamp}_Recording.transcript.vtt", "\n".join(vtt))
    if chat:
        m.sink.write_text(f"{folder}/{stamp}_Recording_newChat.txt", "\n".join(chat) + "\n")
    buf = io.StringIO()
    buf.write("Name (Original Name),User Email,Join Time,Leave Time,Duration (Minutes),Guest\n")
    for p, email, guest in people:
        j = start + dt.timedelta(seconds=r.randint(-120, 180))
        lv = max(j + dt.timedelta(minutes=1), end - dt.timedelta(seconds=r.randint(0, 120)))
        buf.write(f"{p},{email},{j:%m/%d/%Y %I:%M:%S %p},{lv:%m/%d/%Y %I:%M:%S %p},"
                  f"{max(1, int((lv - j).total_seconds() // 60))},{'Yes' if guest else 'No'}\n")
    m.sink.write_text(f"{folder}/participants_{mid}.csv", buf.getvalue())
    return end


def meet_export(m, folder, title, start, people, lines, minutes=30):
    """Google Meet exports: transcript (Docs, saved as .docx) and attendance report .csv."""
    C, r = m.C, m.rng
    d = C.new_doc(f"{title} - Transcript", f"{start:%b %d, %Y}")
    d.add_heading("Attendees", 1)
    d.add_paragraph(", ".join(p for p, _, _ in people))
    d.add_heading("Transcript", 1)
    t = 0.0
    gap = minutes * 60 / max(len(lines), 1)
    for who, text in lines:
        if r.random() < 0.25 or t == 0:
            d.add_paragraph(f"{int(t // 3600):02d}:{int(t % 3600 // 60):02d}:{int(t % 60):02d}").runs[0].bold = True
        d.add_paragraph(f"{who}: {text}")
        t += max(3 + len(text) / 14, gap * r.uniform(0.6, 1.4))
    d.add_paragraph("Meeting ended after " + f"{int(t // 3600):02d}:{int(t % 3600 // 60):02d}:{int(t % 60):02d}")
    C.save_doc(m.sink, f"{folder}/{safe(title, 80)} ({start:%Y-%m-%d %H%M} GMT) - Transcript.docx", d)
    buf = io.StringIO()
    buf.write("First name,Last name,Email,Duration,Time joined,Time exited\n")
    end = start + dt.timedelta(seconds=t + 20)
    for p, email, _ in people:
        first, last = (p.split(" ", 1) + [""])[:2]
        j = start + dt.timedelta(seconds=r.randint(0, 240))
        mins = max(1, int((end - j).total_seconds() // 60))
        buf.write(f"{first},{last},{email},{mins // 60} hr {mins % 60} min,{j:%I:%M %p},{end:%I:%M %p}\n")
    m.sink.write_text(f"{folder}/{safe(title, 80)} - Attendance report.csv", buf.getvalue())


def _who(e):
    return (e["name"], e["email"], False)


def meetings_av(m):
    C, r = m.C, m.rng
    mgrs = [e for e in m.staff if e["role"] in C.HEADS.values()]
    if not mgrs:
        return
    host = mgrs[0]["name"]
    base = f"{m.F['mail']}/{m.y}"
    # weekly operations meeting on Zoom
    for day in range(0, m.days, 7):
        date = m.d0 + dt.timedelta(days=day)
        al = [a for a in m.alerts if (a[0].date() - date).days in range(-7, 0)]
        inc = [i for i in m.incs if (i["opened"].date() - date).days in range(-7, 0)]
        fp = sum(1 for a in al if a[4].startswith("False"))
        soc = next((e for e in mgrs if e["role"] == "SOC Manager"), mgrs[0])
        fld = next((e for e in mgrs if e["role"] == "Field Services Manager"), mgrs[0])
        fin = next((e for e in mgrs if e["role"] == "Finance Manager"), mgrs[0])
        lines = [(host, "Morning everyone, let's start. Numbers from last week first."),
                 (soc["name"], f"We handled {len(al)} alarms, {fp} of them false. That's "
                               f"{round(100 * fp / max(len(al), 1))} percent false, roughly in line with last week."),
                 (soc["name"], f"{len(inc)} incidents, {sum(1 for i in inc if i['pri'] == 'P1')} at P1."
                               + (f" The worst was {inc[0]['cat'].lower()} at {inc[0]['site']['name']}." if inc else "")),
                 (host, "Any SLA misses on the critical alarms?"),
                 (soc["name"], r.choice(["Two on the night crew, both during the shift change. I'll fix the handover overlap.",
                                         "None on critical. A few medium ones slipped past five minutes.",
                                         "One, a video loss that sat in the queue. We changed the routing rule."])),
                 (fld["name"], f"Field team closed {len([w for w in m.workorders if w[1][:10] >= str(date - dt.timedelta(days=7))])} "
                               "work orders. Spare 4K bullet cameras are low, I've raised a purchase order."),
                 (fin["name"], r.choice(["Receivables over sixty days are up. I'll send the list to account managers today.",
                                         "Cash is fine. Two clients are late, both promised payment this week.",
                                         "Expense claims are due Friday, please remind your teams."])),
                 (host, r.choice(["Firmware rollout: where are we?", "How is the night-shift hiring going?",
                                  "Any client escalations I should know about?"])),
                 (r.choice(mgrs)["name"], r.choice(["About seventy percent done, the rest is booked for next week.",
                                                    "Two offers out, one accepted.",
                                                    "One hotel asked for a service credit; we're reviewing the timeline."])),
                 (host, "Good. Actions are in the minutes. Thanks all.")]
        zoom_export(m, f"{base}/Zoom Meetings/{date} 0900 Weekly operations meeting",
                    "Weekly operations meeting", dt.datetime.combine(date, dt.time(9, 0)), [_who(e) for e in mgrs],
                    lines, host, r.randint(30, 50))
    # P1 incident bridge calls on Google Meet
    for i in m.incs:
        if i["pri"] != "P1":
            continue
        cl = m.cl_by[i["site"]["client"]]
        start = i["opened"] + dt.timedelta(minutes=r.randint(6, 15))
        own, an = i["owner"]["name"], i["analyst"]["name"]
        lines = [(own, f"Bridge open for {i['id']}, {i['cat'].lower()} at {i['site']['name']}. {an}, what do we see?"),
                 (an, f"Alarm was '{i['typ']}' at {i['opened']:%H:%M}. I reviewed live and recorded video and confirmed it."),
                 (own, f"Has {cl['contact']} been called?"),
                 (an, "Yes, the duty manager is on the phone with their security team now."),
                 (own, "Preserve the footage from all cameras in that zone, ninety-day evidence hold."),
                 (an, "Hold is set. Export is running."),
                 (own, r.choice(["Do we need police?", "Any sign of a second person?", "Is the recorder healthy?"])),
                 (an, r.choice(["Client says they'll call them.", "No, only one person on all angles.",
                                "Recorder is fine, no gaps in the timeline."])),
                 (own, f"Root cause for now: {i['root'].lower()}. I'll write the incident report. Closing the bridge.")]
        meet_export(m, f"{base}/Google Meet/{start:%Y-%m-%d} Incident bridge {i['id']}", f"Incident bridge {i['id']}",
                    start, [_who(i["owner"]), _who(i["analyst"]), (cl["contact"], f"security@{cl['domain']}", True)], lines,
                    r.randint(20, 70))
    # monthly change advisory board on Google Meet
    first_tue = m.d0 + dt.timedelta(days=(1 - m.d0.weekday()) % 7)
    eng = [e for e in m.staff if e["role"] in ("Network Engineer", "Cloud Engineer")][:4]
    changes = r.sample(["Recorder firmware upgrade on ten sites", "Firewall rule for new client VPN",
                        "Move cloud video storage to a cheaper tier", "Rotate recorder admin passwords",
                        "Patch analytics servers", "Replace core switch at the SOC", "Enable MFA on the VMS web client",
                        "Increase retention to 60 days for one hotel"], 3)
    lines = [(host, "Change board. Three changes on the list.")]
    for c in changes:
        who = (r.choice(eng) if eng else mgrs[0])["name"]
        lines += [(who, f"{c}. Risk is low to medium; rollback plan is tested. Window is Sunday 02:00 to 04:00."),
                  (host, r.choice(["Any client impact?", "Who is on call?", "Do we notify clients?"])),
                  (who, r.choice(["Live view drops for about two minutes per site.", "I am, with the night shift lead.",
                                  "Yes, the notice goes out 72 hours before."])),
                  (host, r.choice(["Approved.", "Approved, go ahead.", "Approved with a smaller first batch."]))]
    meet_export(m, f"{base}/Google Meet/{first_tue} Change advisory board", "Change advisory board",
                dt.datetime.combine(first_tue, dt.time(14, 0)), [_who(e) for e in (mgrs + eng)], lines, r.randint(30, 45))
    # quarterly client service reviews on Zoom
    if m.ym[1] % 3 == 0:
        ams = [e for e in m.staff if e["role"] == "Account Manager"] or mgrs
        for c in m.clients:
            day = 10 + c["id"] % 15
            start = dt.datetime.combine(m.d0 + dt.timedelta(days=min(day, m.days - 1)), dt.time(15, 0))
            sids = set(c["sites"])
            al = [a for a in m.alerts if a[1]["id"] in sids]
            inc = [i for i in m.incs if i["site"]["id"] in sids]
            am = ams[c["id"] % len(ams)]
            lines = [(am["name"], f"Thanks for joining, {c['contact'].split()[0]}. This is the quarterly service review."),
                     (am["name"], f"This month alone we handled {len(al)} alarms across your sites and {len(inc)} incidents."),
                     (c["contact"], r.choice(["The night team has been quick, our guards noticed.",
                                              "We had complaints about false alarms in the lobby.",
                                              "Can we get footage exports faster for the police?"])),
                     (am["name"], r.choice(["We'll retune the lobby zones this week.",
                                            "Exports for police now go out within two hours.",
                                            "Thanks, I'll pass that to the SOC."])),
                     (c["contact"], "What about the older cameras?"),
                     (am["name"], "Twelve are past five years. I'll send a replacement quote with the report."),
                     (c["contact"], "Good. Send me the deck after the call.")]
            zoom_export(m, f"{base}/Zoom Meetings/{start:%Y-%m-%d} 1500 Quarterly service review {safe(c['name'], 50)}",
                        "Quarterly service review", start,
                        [_who(am), (c["contact"], f"security@{c['domain']}", True)], lines, am["name"], r.randint(35, 55))


# ---------------------------------------------------------------- IT SOPs

IT_SOPS = {
    "SOP-IT-01 New user onboarding": [
        ("Trigger", ["HR raises an onboarding ticket at least five working days before the start date."]),
        ("Steps", ["Create the directory account from the role template; never copy another user.",
                   "Create the mailbox and add the user to the department distribution list.",
                   "Issue a laptop from stock; confirm disk encryption and endpoint protection are active.",
                   "Enrol MFA on the company phone with the user present on day one.",
                   "Grant application access listed in the role matrix only (VMS, ticketing, alarm platform).",
                   "Record asset tags against the user in the asset register.",
                   "Close the ticket with the checklist attached."])],
    "SOP-IT-02 Leaver offboarding": [
        ("Trigger", ["HR ticket with the last working day; urgent if the departure is involuntary."]),
        ("Steps", ["Disable the directory account at 17:00 on the last day (immediately if involuntary).",
                   "Revoke MFA tokens, VPN profile and all VMS and recorder accounts.",
                   "Rotate any shared vault entries the user could see.",
                   "Convert the mailbox to shared and give the manager access for 90 days.",
                   "Collect and wipe the laptop and phone; update the asset register.",
                   "Ask each client site to deactivate the user's access card; record confirmations."])],
    "SOP-IT-03 Password reset and account unlock": [
        ("Steps", ["Verify identity with a call-back to the number in the HR system, never the number given.",
                   "Reset with a one-time password and force change at next logon.",
                   "For VMS client users, confirm the request with the client's named security contact.",
                   "Log the ticket with the verification method used."])],
    "SOP-IT-04 Patch management": [
        ("Schedule", ["Workstations and servers: monthly, within 14 days of vendor release.",
                      "Critical vulnerabilities on the CISA KEV list: within 72 hours.",
                      "Recorders and cameras: quarterly firmware window, sites in batches of ten."]),
        ("Steps", ["Check vendor notes and test on the lab recorder first.", "Raise a change for the change board.",
                   "Notify affected clients 72 hours ahead.", "Apply, then confirm recording and live view on every channel.",
                   "Roll back if more than 2% of channels fail to recover in 15 minutes."])],
    "SOP-IT-05 Backup and restore": [
        ("Scope", ["Ticketing database, configuration exports of recorders and switches, file shares, finance system."]),
        ("Steps", ["Nightly backups to a separate cloud account with immutability for 30 days.",
                   "Weekly test restore of one system chosen at random; record the result.",
                   "Quarterly full restore drill of the ticketing system to a clean environment.",
                   "Report failed jobs to the Cloud Engineer on call within one hour."])],
    "SOP-IT-06 Change management": [
        ("Steps", ["Raise a change with description, risk, rollback plan and window.",
                   "Standard changes (pre-approved list) go ahead after peer review.",
                   "Normal changes go to the weekly change board; emergency changes need the Operations Director.",
                   "After the change, record the outcome and any incident links."])],
    "SOP-IT-07 Recorder and camera firmware upgrade": [
        ("Steps", ["Export the recorder configuration before any change.", "Confirm UPS and network health on site.",
                   "Upgrade the recorder, then cameras in groups of eight.",
                   "Check each channel: live view, recording, analytics rules, time sync.",
                   "Update the asset register with new firmware versions."])],
    "SOP-IT-08 Major incident (P1) bridge": [
        ("Steps", ["Shift lead opens a Google Meet bridge within 10 minutes of a P1.",
                   "Analyst reports what was seen and when; shift lead owns decisions.",
                   "Call the client's duty manager; record the time of the call.",
                   "Set a 90-day evidence hold on all cameras in the zone.",
                   "Post updates every 15 minutes in the incident channel.",
                   "Write the incident report within 24 hours."])],
    "SOP-IT-09 Footage export for police and insurers": [
        ("Steps", ["Accept requests only from the client's named contacts or with a police reference.",
                   "Export to an encrypted archive with a SHA-256 hash recorded in the ticket.",
                   "Share by expiring link; never by USB unless the police collect in person.",
                   "Keep the chain-of-custody form with the ticket."])],
    "SOP-IT-10 Remote access and VPN": [
        ("Steps", ["All remote access to recorders goes through the company VPN with MFA.",
                   "Recorder web ports are never exposed to the internet; quarterly external scan confirms this.",
                   "Vendor access is time-limited and supervised.", "VPN logs are kept for 12 months."])],
    "SOP-IT-11 Network switch replacement": [
        ("Steps", ["Pull the running configuration from the backup system.", "Label every patch lead before removal.",
                   "Install the new switch, load the configuration, check PoE budget.",
                   "Confirm every camera returns online in the VMS; record downtime per channel."])],
    "SOP-IT-12 Vulnerability management": [
        ("Steps", ["Weekly authenticated scan of servers and workstations; monthly scan of recorders.",
                   "Match findings to the KEV catalogue; KEV items are fixed first.",
                   "Track exceptions with an owner and expiry date.", "Report open critical items to the board each quarter."])],
}


def it_sops(w, sink, C):
    r = random.Random(C.SEED + 11)
    owner = next((e["name"] for e in w.staff if e["role"] == "Operations Director"), w.staff[0]["name"])
    for title, sections in IT_SOPS.items():
        d = C.new_doc(title, f"IT Standard Operating Procedure | Version {r.randint(1, 4)}.{r.randint(0, 9)}")
        C.doc_table(d, ["Owner", "Applies to", "Review"], [[owner, "IT Support, Network, Cloud, SOC", "Annual"]])
        for h, steps in sections:
            d.add_heading(h, 1)
            for i, s in enumerate(steps, 1):
                d.add_paragraph(f"{i}. {s}")
        d.add_heading("Records", 1)
        d.add_paragraph("Every run of this procedure is logged in the ticketing system with the checklist attached.")
        C.save_doc(sink, f"{C.F['pol']}/IT SOPs/{title}.docx", d)
