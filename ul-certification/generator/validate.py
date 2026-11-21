# -*- coding: utf-8 -*-
"""usage: python3 /tmp/deck/work/validate.py <unit>   e.g. rows03 | cards02 | main_rows01 | main_other"""
import json, glob, sys, re

unit = sys.argv[1]
inp = json.load(open(f"/tmp/deck/work/in/{unit}.json", encoding="utf-8"))
typ = inp["type"]
want = {it["k"]: it for it in inp["items"]}
got = {}
errs = []
for fp in sorted(glob.glob(f"/tmp/deck/work/out/{unit}.part*.json")):
    try:
        arr = json.load(open(fp, encoding="utf-8"))
    except Exception as e:
        errs.append(f"{fp}: invalid JSON {e}")
        continue
    for e in arr:
        k = e.get("k")
        if k in got:
            errs.append(f"duplicate {k}")
        got[k] = e

LIM = {"req": 125, "why": 105, "verify": 105, "pitfall": 105, "cite": 100}
BAN = re.compile(r"agent-tools|/tmp/|\.txt|subagent|\n")

if typ in ("rows", "main_rows"):
    for k, e in got.items():
        if k not in want:
            errs.append(f"unknown key {k}")
            continue
        for f, lim in LIM.items():
            v = e.get(f)
            if not isinstance(v, str) or not v.strip():
                errs.append(f"{k}: missing {f}")
            elif len(v) > lim:
                errs.append(f"{k}: {f} too long {len(v)}>{lim}")
            elif BAN.search(v):
                errs.append(f"{k}: {f} contains banned text")
        if set(e) - set(LIM) - {"k"}:
            errs.append(f"{k}: extra fields {set(e) - set(LIM) - {'k'}}")
else:
    for k, e in got.items():
        if k not in want:
            errs.append(f"unknown key {k}")
            continue
        sl = e.get("slides")
        if not isinstance(sl, list) or not (1 <= len(sl) <= 3):
            errs.append(f"{k}: need 1-3 slides")
            continue
        for i, s in enumerate(sl):
            if s.get("kind") != "cards":
                errs.append(f"{k}[{i}]: kind must be cards")
            if len(s.get("title", "")) > 34:
                errs.append(f"{k}[{i}]: title too long")
            cs = s.get("cards", [])
            if not (1 <= len(cs) <= 2):
                errs.append(f"{k}[{i}]: 1-2 cards")
            for c in cs:
                if c.get("color") not in ("teal", "orange", "red", "green"):
                    errs.append(f"{k}[{i}]: bad color")
                if not (1 <= len(c.get("items", [])) <= 5):
                    errs.append(f"{k}[{i}]: 1-5 items per card")
                for it in c.get("items", []):
                    if len(it) > 150:
                        errs.append(f"{k}[{i}]: item too long {len(it)}>150")
                    if BAN.search(it):
                        errs.append(f"{k}[{i}]: banned text")
missing = [k for k in want if k not in got]
print(f"unit={unit} want={len(want)} got={len(got)} missing={len(missing)} errors={len(errs)}")
for e in errs[:40]:
    print("ERR", e)
if missing:
    print("MISSING (first 10):", missing[:10])
