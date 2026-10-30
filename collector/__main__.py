"""python -m collector --sources kev,nvd,... --live

Writes into OUT, ships to the rclone remote after each source, keeps state in STATE.
Logs show counts only."""
import argparse
import os
import time
import traceback

from . import company, github, kaggle, library, named, public, studio, techdocs
from .common import Sink, State

EXTRA = {**studio.SOURCES, **named.SOURCES, **library.SOURCES, **techdocs.SOURCES, **kaggle.SOURCES, **github.SOURCES}
ALL = list(public.SOURCES) + list(EXTRA) + ["company"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sources", default=",".join(ALL))
    ap.add_argument("--out", default="out")
    ap.add_argument("--state", default="state/state.json")
    ap.add_argument("--remote", default=os.environ.get("REMOTE", "pax:Studio Operations Research"))
    ap.add_argument("--live", action="store_true")
    ap.add_argument("--keep", action="store_true", help="dry run: keep files in OUT for inspection")
    ap.add_argument("--budget-gb", type=float, default=20)
    ap.add_argument("--minutes", type=float, default=320)
    ap.add_argument("--months", type=int, default=24, help="company months per run")
    ap.add_argument("--scale", type=float, default=1.0, help="company size multiplier")
    ap.add_argument("--detail", type=float, default=1.0, help="telemetry density (0 = none)")
    ap.add_argument("--nvd-months", type=int, default=24)
    a = ap.parse_args()
    st = State(a.state)
    sink = Sink(a.out, a.remote, a.live, st, int(a.budget_gb * 2 ** 30), time.time() + a.minutes * 60, a.keep)
    srcs = [x.strip() for x in a.sources.split(",") if x.strip()]
    if "company" in srcs:  # long-running sources last so fixed ones always get a turn
        srcs = [x for x in srcs if x != "company"] + ["company"]
    failed = []
    for i, name in enumerate(srcs):
        # fair share of the time left: no single source can starve the ones after it
        left = sink.deadline - time.time()
        sink.slice_end = time.time() + max(15 * 60, left / (len(srcs) - i))
        if sink.full():
            print(f"{name}: skipped (time or size budget reached)")
            continue
        t0 = time.time()
        try:
            if name == "company":
                sec = st.section("company")
                sec.setdefault("scale", a.scale)  # scale is fixed once the history starts
                n = company.build(sink, sec, a.months, sec["scale"], a.detail)
                unit = "months"
            elif name in EXTRA:
                n, unit = EXTRA[name](sink, st.section("public")), "items"
            else:
                n, unit = public.SOURCES[name](sink, st.section("public")), "items"
            sink.flush()
            print(f"{name}: {n} {unit} in {time.time() - t0:.0f}s; total {sink.files} files, "
                  f"{sink.sent_bytes / 2 ** 20:.0f} MB")
        except Exception:
            failed.append(name)
            print(f"{name}: FAILED")
            traceback.print_exc()
            try:
                sink.flush()
            except Exception:
                traceback.print_exc()
    print(f"done: {sink.files} files, {sink.sent_bytes / 2 ** 30:.2f} GB, live={a.live}")
    if failed:
        raise SystemExit(f"failed sources: {', '.join(failed)}")


if __name__ == "__main__":
    main()
