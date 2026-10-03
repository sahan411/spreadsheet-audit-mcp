"""Tests for xlsx_audit.py: builds a deliberately broken workbook and a clean one, then checks each rule fires (and does not fire on the clean file)."""
import os, sys, tempfile, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import xlsx_audit as A
from fixtures import broken, clean, extlink

TMP = tempfile.mkdtemp()


def rules(res):
    return {f["rule"] for f in res["findings"]}


def main():
    ok = fail = 0

    def check(name, cond, extra=""):
        nonlocal ok, fail
        ok, fail = (ok + 1, fail) if cond else (ok, fail + 1)
        print("PASS" if cond else "FAIL", name, extra)

    b, c, e = (os.path.join(TMP, n) for n in ("broken.xlsx", "clean.xlsx", "ext.xlsx"))
    broken(b); clean(c); extlink(e)
    rb = A.audit(b)
    got = rules(rb)
    for rule in ("TYPED-OVER", "PATTERN-BREAK", "MAGIC-NUMBER", "LOOKUP-APPROX", "VOLATILE", "HIDDEN-SHEET", "CALC-ERROR"):
        check(f"broken: {rule} detected", rule in got, "" if rule in got else f"got {sorted(got)}; note={rb['recalculation_note']}")
    cells = {(f["rule"], f["cell"]) for f in rb["findings"]}
    check("TYPED-OVER points at D5", ("TYPED-OVER", "D5") in cells)
    check("PATTERN-BREAK points at E6", ("PATTERN-BREAK", "E6") in cells)
    check("exact VLOOKUP (F3) not flagged", ("LOOKUP-APPROX", "F3") not in cells)
    check("approx VLOOKUP (F2) flagged", ("LOOKUP-APPROX", "F2") in cells)
    check("CALC-ERROR points at G2", ("CALC-ERROR", "G2") in cells)
    rc = A.audit(c)
    check("clean workbook has no HIGH or MEDIUM findings", not any(f["severity"] in ("HIGH", "MEDIUM") for f in rc["findings"]), str(rc["counts"]))
    re_ = A.audit(e, recalc=False)
    check("external link detected", "EXT-LINK" in rules(re_), str(rules(re_)))
    print(f"\n{ok} passed, {fail} failed")
    shutil.rmtree(TMP, ignore_errors=True)
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
