"""Aggregate the UKE PlanCheck review logs (Jul 2023 - Dec 2024) without exporting identifiers.

Usage: python -I aggregate_plancheck.py <folder with PatientData_PlanCheck.csv, FoundFixes.csv, finalerrors.csv, settings.json> <out dir>

Writes only counts. Patient, course and plan identifiers and free-text messages are never written;
messages are reduced to a template (digits and quoted names removed) only to label error codes.

Definitions (from the PlanCheck viewer, dataframe.py):
  plan          = patient/course/plan with >= 2 PlanCheck runs by >= 2 different users
  major/minor   = PlanCheck class "1 - Deviation" / "2 - Variation"
  resolved      = finding present in one run and absent from the next run of the same plan (FoundFixes)
  open at last  = finding still present in the last run (finalerrors)
  codes switched off in settings.json are not counted (they were excluded in the viewer)
"""
import json, os, re, sys, collections
import pandas as pd

SRC, OUT = sys.argv[1], sys.argv[2]
os.makedirs(OUT, exist_ok=True)

CATEGORY = {  # error code -> category (assigned from the message template; checked by hand)
    2: "Treatment unit vs. scheduled appointment", 4: "Treatment unit vs. scheduled appointment",
    96: "Prescription and reference points", 108: "Prescription and reference points", 109: "Prescription and reference points",
    60: "Field naming and setup fields", 62: "Field naming and setup fields", 64: "Field naming and setup fields",
    66: "Field naming and setup fields", 68: "Field naming and setup fields", 91: "Field naming and setup fields",
    93: "Field naming and setup fields", 54: "Field naming and setup fields",
    106: "Dose calculation grid", 115: "Dose calculation grid", 123: "Dose calculation grid",
    23: "Dose maximum outside target", 27: "Dose maximum level", 30: "Dose maximum level",
    51: "Couch and collision geometry", 113: "Couch and collision geometry",
    83: "MLC and delivery settings", 84: "MLC and delivery settings", 55: "MLC and delivery settings",
    57: "MLC and delivery settings", 71: "MLC and delivery settings", 90: "MLC and delivery settings",
    20: "CT calibration", 7: "Motion management (gating)",
    121: "Cardiac implant", 122: "Cardiac implant",
    25: "Isocentre, origin and laterality", 79: "Isocentre, origin and laterality", 81: "Isocentre, origin and laterality",
    88: "Isocentre, origin and laterality", 98: "Isocentre, origin and laterality", 100: "Isocentre, origin and laterality",
    102: "Isocentre, origin and laterality",
    24: "Target structure", 116: "Course and plan identity", 41: "Course and plan identity", 114: "Couch and collision geometry",
}


def read(name):
    return pd.read_csv(os.path.join(SRC, name), sep=";", encoding="utf-8")


P = read("PatientData_PlanCheck.csv")
F = read("FoundFixes.csv")
E = read("finalerrors.csv")
excluded = sorted(int(k) for k, v in json.load(open(os.path.join(SRC, "settings.json"))).items() if v)
P["t0"] = pd.to_datetime(P["Erstes Öffnen"], format="%Y.%m.%d %H:%M")
F["t"] = pd.to_datetime(F["Zeit"], format="%Y.%m.%d %H:%M", errors="coerce")
P["plan"] = P["Patienten-ID"].astype(str) + "|" + P["Kurs"].astype(str) + "|" + P["Plan"].astype(str)
for D in (F, E):
    D["major"] = D["Art"].str.startswith("1")
    D["cat"] = D["Fehlercode"].map(CATEGORY).fillna("Other")
    D["plan"] = D["PatientID"].astype(str) + "|" + D["Course"].astype(str) + "|" + D["PlanID"].astype(str)

# FoundFixes is written before the viewer drops plans with a single user; keep only included plans
F_all = len(F)
F = F[F.plan.isin(set(P.plan))].copy()
E_all = len(E)
E = E[E.plan.isin(set(P.plan))].copy()


def tmpl(s):
    s = re.sub(r"\d+([.,]\d+)?", "#", str(s))
    s = re.sub(r"'[^']*'|\"[^\"]*\"", "'…'", s)
    return s[:100]


n = len(P)
summary = {
    "period_first_check": str(P.t0.min().date()), "period_last_check": str(P.t0.max().date()),
    "plans": n, "patients": int(P["Patienten-ID"].nunique()),
    "runs_per_plan_median": float(P["Anzahl"].median()),
    "runs_per_plan_q1": float(P["Anzahl"].quantile(.25)), "runs_per_plan_q3": float(P["Anzahl"].quantile(.75)),
    "runs_total": int(P["Anzahl"].sum()),
    "users_per_plan": {int(k): int(v) for k, v in P["User#"].value_counts().sort_index().items()},
    "plans_with_major_any_run_viewer": int((P["Dev [Max]"] > 0).sum()),
    "plans_with_minor_any_run_viewer": int((P["Var [Max]"] > 0).sum()),
    "plans_with_major_decrease": int((P["Dev [∆]"] > 0).sum()),
    "plans_with_minor_decrease": int((P["Var [∆]"] > 0).sum()),
    "resolved_rows_before_plan_filter": F_all, "open_rows_before_plan_filter": E_all,
    "resolved_findings": len(F), "resolved_major": int(F.major.sum()), "resolved_minor": int((~F.major).sum()),
    "plans_with_resolved_major": int(F[F.major].plan.nunique()), "plans_with_resolved_any": int(F.plan.nunique()),
    "open_at_last_findings": len(E), "open_at_last_major": int(E.major.sum()), "open_at_last_minor": int((~E.major).sum()),
    "plans_with_open_major": int(E[E.major].plan.nunique()),
    "codes_excluded_in_viewer": excluded,
    "codes_observed": int(pd.concat([F.Fehlercode, E.Fehlercode]).nunique()),
    "plans_by_year": {int(k): int(v) for k, v in P.t0.dt.year.value_counts().sort_index().items()},
}
json.dump(summary, open(os.path.join(OUT, "summary.json"), "w"), indent=1)

rows = []
for cat in sorted(set(F.cat) | set(E.cat)):
    f, e = F[F.cat == cat], E[E.cat == cat]
    rows.append({"category": cat,
                 "resolved_major": int(f.major.sum()), "resolved_minor": int((~f.major).sum()),
                 "plans_resolved": int(f.plan.nunique()),
                 "open_major": int(e.major.sum()), "open_minor": int((~e.major).sum()),
                 "plans_open": int(e.plan.nunique()),
                 "plans_any_run": len(set(f.plan) | set(e.plan)),
                 "plans_any_run_major": len(set(f[f.major].plan) | set(e[e.major].plan)),
                 "codes": ",".join(str(c) for c in sorted(set(f.Fehlercode) | set(e.Fehlercode)))})
cat_df = pd.DataFrame(rows).sort_values(["plans_any_run", "plans_resolved"], ascending=False)
key = ["plan", "Fehlercode", "Art", "Beschreibung"]
summary["resolved_distinct_findings"] = len(F.drop_duplicates(key))
summary["resolved_present_again_at_last_run"] = len(F.merge(E[key].drop_duplicates(), on=key).drop_duplicates(key))
summary["plans_with_open_any"] = int(E.plan.nunique())
summary["plans_any_finding_from_rows"] = len(set(F.plan) | set(E.plan))
summary["plans_any_major_from_rows"] = len(set(F[F.major].plan) | set(E[E.major].plan))
json.dump(summary, open(os.path.join(OUT, "summary.json"), "w"), indent=1)
cat_df.to_csv(os.path.join(OUT, "by_category.csv"), index=False)

code_rows = []
for code in sorted(set(F.Fehlercode) | set(E.Fehlercode)):
    f, e = F[F.Fehlercode == code], E[E.Fehlercode == code]
    allmsg = pd.concat([f.Beschreibung, e.Beschreibung])
    first = f.t.min()
    code_rows.append({"code": int(code), "category": CATEGORY.get(int(code), "Other"),
                      "class": "major" if pd.concat([f.major, e.major]).mean() >= .5 else "minor",
                      "resolved": len(f), "open_at_last": len(e),
                      "first_resolved_month": "" if pd.isna(first) else first.strftime("%Y-%m"),
                      "message_template": collections.Counter(allmsg.map(tmpl)).most_common(1)[0][0]})
pd.DataFrame(code_rows).to_csv(os.path.join(OUT, "by_code.csv"), index=False)

month = F.dropna(subset=["t"]).assign(m=lambda d: d.t.dt.to_period("M").astype(str))
mon = month.groupby(["m", "major"]).size().unstack(fill_value=0).rename(columns={True: "resolved_major", False: "resolved_minor"})
mon = mon.join(P.assign(m=P.t0.dt.to_period("M").astype(str)).groupby("m").size().rename("plans_first_checked"), how="outer").fillna(0).astype(int)
mon.to_csv(os.path.join(OUT, "by_month.csv"))
print(json.dumps(summary, indent=1))
print(cat_df.to_string(index=False))
