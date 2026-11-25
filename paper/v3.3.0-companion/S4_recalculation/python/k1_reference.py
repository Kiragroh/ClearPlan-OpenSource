"""Independent K1 reference calculation for the ClearPlan phantom manuscript.

Recomputes DVH metrics, Paddick CI, GI, HI, aperture areas and PAM from the primitive
ESAPI export (ClearPlanK1ReferenceExport) without any ClearPlan code, then compares
them with the ClearPlan baseline review log under tolerances fixed in tolerances.json.

Usage:
  python k1_reference.py compare --export <k1-*.json> [--export ...] --label-fractions single=5 double=40
      --sqlite <native_review_log.sqlite> --runs <native_verification.json>
      --tolerances <tolerances.json> --mlc <mlc_geometry.json> [--gui <gui_values.csv>] --out <dir>
"""
import argparse
import csv
import json
import math
import os
import sqlite3
import sys

import numpy as np
import shapely
from shapely.geometry import Polygon, box
from shapely.ops import unary_union


# ---------------------------------------------------------------- export and dose

def load_export(path):
    with open(path, encoding="utf8") as f:
        exp = json.load(f)
    d = exp["dose"]
    nx, ny, nz = d["size"]
    raw = np.fromfile(os.path.join(os.path.dirname(path), d["file"]), dtype="<i4")
    if raw.size != nx * ny * nz:
        raise ValueError("dose file size does not match the grid")
    exp["_dose"] = DoseGrid(raw.reshape(nz, ny, nx) * d["gy_per_raw"] + d["gy_offset"],
                            d["origin"], d["res_mm"], d["x_dir"], d["y_dir"], d["z_dir"])
    return exp


class DoseGrid:
    def __init__(self, gy, origin, res, x_dir, y_dir, z_dir):
        for axis, expected in ((x_dir, (1, 0, 0)), (y_dir, (0, 1, 0)), (z_dir, (0, 0, 1))):
            if np.abs(np.subtract(axis, expected)).max() > 1e-6:
                raise ValueError("only axis-aligned dose grids are supported")
        self.gy = np.asarray(gy, dtype=np.float64)
        self.origin = np.asarray(origin, dtype=np.float64)
        self.res = np.asarray(res, dtype=np.float64)

    def at(self, points):
        """Trilinear dose at Nx3 points (mm); zero outside the grid, with an outside count."""
        idx = (np.asarray(points) - self.origin) / self.res  # voxel-centre index coordinates
        nz, ny, nx = self.gy.shape
        size = np.array([nx, ny, nz])
        inside = np.all((idx >= 0) & (idx <= size - 1), axis=1)
        out = np.zeros(len(idx))
        i = idx[inside]
        i0 = np.minimum(np.floor(i).astype(int), size - 2)
        f = i - i0
        x0, y0, z0 = i0[:, 0], i0[:, 1], i0[:, 2]
        fx, fy, fz = f[:, 0], f[:, 1], f[:, 2]
        g = self.gy
        c00 = g[z0, y0, x0] * (1 - fx) + g[z0, y0, x0 + 1] * fx
        c10 = g[z0, y0 + 1, x0] * (1 - fx) + g[z0, y0 + 1, x0 + 1] * fx
        c01 = g[z0 + 1, y0, x0] * (1 - fx) + g[z0 + 1, y0, x0 + 1] * fx
        c11 = g[z0 + 1, y0 + 1, x0] * (1 - fx) + g[z0 + 1, y0 + 1, x0 + 1] * fx
        out[inside] = (c00 * (1 - fy) + c10 * fy) * (1 - fz) + (c01 * (1 - fy) + c11 * fy) * fz
        return out, int((~inside).sum())


# ---------------------------------------------------------------- structures

def planes(structure):
    """Contours grouped by image plane: {plane: (z, [Nx2 arrays])}."""
    out = {}
    for c in structure["contours"]:
        xyz = np.asarray(c["xyz"], dtype=np.float64)
        z, polys = out.setdefault(c["plane"], (float(xyz[:, 2].mean()), []))
        polys.append(xyz[:, :2])
    return out


def sample_structure(structure, z_res, inplane, z_sub, dose=None, min_dose=None):
    """Sample points inside the structure (even-odd over all contours of a plane).

    Each contour plane is a slab of one image slice thickness centred on the plane.
    Returns (doses or points, weight per sample in cm3, points outside the dose grid).
    """
    w = inplane * inplane * z_res / z_sub / 1000.0
    offsets = (np.arange(z_sub) + 0.5) / z_sub - 0.5
    doses, outside = [], 0
    pts_all = []
    for z, polys in planes(structure).values():
        allxy = np.vstack(polys)
        x = np.arange(math.floor(allxy[:, 0].min() / inplane) * inplane + inplane / 2, allxy[:, 0].max(), inplane)
        y = np.arange(math.floor(allxy[:, 1].min() / inplane) * inplane + inplane / 2, allxy[:, 1].max(), inplane)
        gx, gy = np.meshgrid(x, y)
        gx, gy = gx.ravel(), gy.ravel()
        inside = np.zeros(gx.size, dtype=bool)
        for p in polys:
            inside ^= shapely.contains_xy(Polygon(p).buffer(0), gx, gy)
        gx, gy = gx[inside], gy[inside]
        for o in offsets:
            pts = np.column_stack([gx, gy, np.full(gx.size, z + o * z_res)])
            if dose is None:
                pts_all.append(pts)
                continue
            d, n_out = dose.at(pts)
            outside += n_out
            doses.append(d if min_dose is None else d[d >= min_dose])
    if dose is None:
        return np.vstack(pts_all), w, 0
    return (np.concatenate(doses) if doses else np.zeros(0)), w, outside


def dvh_metrics(d, w, rx):
    """Volume-weighted metrics of equally weighted samples."""
    if d.size == 0:
        return None
    s = np.sort(d)[::-1]
    cum = np.arange(1, s.size + 1) / s.size

    def dose_at(volume_fraction):
        return float(s[min(np.searchsorted(cum, volume_fraction), s.size - 1)])

    return {
        "volume_cm3": d.size * w, "mean_gy": float(d.mean()), "max_gy": float(d.max()),
        "d98_gy": dose_at(0.98), "d2_gy": dose_at(0.02),
        "v_rx_cm3": float((d >= rx).sum() * w), "v20gy_cm3": float((d >= 20).sum() * w),
        "cv20gy_cm3": float((d < 20).sum() * w),
    }


# ---------------------------------------------------------------- apertures and PAM

def aperture(cp, model):
    """Effective aperture polygon (mm at isocentre, collimator frame) of one control point."""
    leaves = np.asarray(cp["leaves"], dtype=np.float64)
    result = None
    for layer in model["layers"]:
        a, b = layer["leaf_indices"]
        bounds = layer["boundaries_mm"]
        strips = [box(leaves[0, i], bounds[i - a], leaves[1, i], bounds[i - a + 1])
                  for i in range(a, b) if leaves[1, i] > leaves[0, i]]
        opening = unary_union(strips) if strips else Polygon()
        result = opening if result is None else result.intersection(opening)
    if model["jaws"] == "physical":
        x1, y1, x2, y2 = cp["jaws_x1_y1_x2_y2"]
        result = result.intersection(box(x1, y1, x2, y2))
    return result


def mu_weights(beam):
    """Trapezoidal MU per control point, as ClearPlan's stated definition."""
    cw = np.array([cp["meterset_weight"] for cp in beam["control_points"]], dtype=np.float64)
    inc = beam["meterset_mu"] * np.diff(cw) / cw[-1]
    w = np.zeros(cw.size)
    w[1:] += inc / 2
    w[:-1] += inc / 2
    return w


def bev_axes(source, iso, collimator_deg, sign):
    """Collimator X and Y axes for HFS, couch 0 (IEC Y_f = DICOM +z)."""
    d = np.asarray(iso) - np.asarray(source)
    sad = np.linalg.norm(d)
    d = d / sad
    yg = np.array([0.0, 0.0, 1.0])
    xg = np.cross(d, yg)
    xg /= np.linalg.norm(xg)
    phi = math.radians(collimator_deg)
    xc = math.cos(phi) * xg + sign * math.sin(phi) * yg
    yc = -sign * math.sin(phi) * xg + math.cos(phi) * yg
    return d, sad, xc, yc


def project(points, source, d, sad, xc, yc):
    v = np.asarray(points) - np.asarray(source)
    q = v * (sad / (v @ d))[:, None]
    return np.column_stack([q @ xc, q @ yc])


def silhouette(structure, z_res, source, iso, collimator, sign, mirror=False):
    """Divergent BEV silhouette: per slab, union of both end faces and the swept edges."""
    d, sad, xc, yc = bev_axes(source, iso, collimator, sign)
    pieces = []
    for z, polys in planes(structure).values():
        for p in polys:
            p = np.asarray(Polygon(p).simplify(0.1).exterior.coords)[:-1]
            lo = project(np.column_stack([p, np.full(len(p), z - z_res / 2)]), source, d, sad, xc, yc)
            hi = project(np.column_stack([p, np.full(len(p), z + z_res / 2)]), source, d, sad, xc, yc)
            if mirror:
                lo[:, 0] *= -1
                hi[:, 0] *= -1
            pieces += [Polygon(lo).buffer(0), Polygon(hi).buffer(0)]
            n = len(p)
            pieces += [Polygon([lo[i], lo[(i + 1) % n], hi[(i + 1) % n], hi[i]]).buffer(0) for i in range(n)]
    return unary_union(pieces)


def native_outline(beam, structure_id):
    for o in beam.get("native_bev_outlines", []):
        if o["structure"] == structure_id and o["polygons"]:
            polys = [Polygon(np.asarray(p).reshape(-1, 2)).buffer(0) for p in o["polygons"]]
            out = Polygon()
            for p in polys:  # even-odd, as for contours
                out = out.symmetric_difference(p)
            return out
    return None


# Collimator rotation follows the IEC 61217 right-hand rule (positive counter-clockwise seen from
# the source), the same rule as the gantry axes that the native check below confirms.
COLLIMATOR_SIGN = 1


def calibrate_orientation(beam, target, z_res):
    """Check the gantry-frame axes against the native CP0 silhouette.

    Beam.GetStructureOutlines(inBEV) returns the outline without collimator rotation
    (found in the 3 Nov 2026 export), so the comparison runs at collimator 0 and can only
    decide mirroring; the collimator sign is fixed by COLLIMATOR_SIGN.
    """
    native = native_outline(beam, target["id"])
    if native is None or native.area <= 0:
        return None, {}
    cp0 = beam["control_points"][0]
    scores = {}
    for mirror in (False, True):
        sil = silhouette(target, z_res, cp0["source"], beam["isocenter"], 0.0, COLLIMATOR_SIGN, mirror)
        scores[mirror] = sil.symmetric_difference(native).area / native.union(sil).area
    best = min(scores, key=scores.get)
    return (COLLIMATOR_SIGN, best), {("mirrored" if m else "direct"): v for m, v in scores.items()}


def beam_metrics(beam, target, z_res, mlc, orientation):
    model = mlc["models"].get(beam["mlc_model"])
    if model is None:
        return {"error": "no leaf geometry for MLC model %r" % beam["mlc_model"]}
    w = mu_weights(beam)
    areas, blocked = [], []
    for cp, wi in zip(beam["control_points"], w):
        ap = aperture(cp, model)
        areas.append(ap.area / 100.0)
        if orientation is None or wi <= 0:
            blocked.append(np.nan)
            continue
        if abs(cp["couch"]) > 1e-6:
            return {"error": "PAM reference supports couch 0 only"}
        sil = silhouette(target, z_res, cp["source"], beam["isocenter"], cp["collimator"], *orientation)
        blocked.append(1 - sil.intersection(ap).area / sil.area if sil.area > 0 else np.nan)
    areas, blocked = np.array(areas), np.array(blocked)
    mu = beam["meterset_mu"]
    out = {"total_mu": mu, "mean_aperture_cm2": float((w * areas).sum() / mu),
           "small_aperture_fraction": float(w[areas < 4].sum() / mu)}
    active = w > 0
    out["pam"] = float((w[active] * blocked[active]).sum() / mu) if not np.isnan(blocked[active]).any() else None
    return out


# ---------------------------------------------------------------- plan reference

def reference(exp, tol, mlc, target_id, goal_structures):
    dose, rx, z_res = exp["_dose"], exp["rx_gy"], exp["image"]["z_res_mm"]
    smp = tol["sampling"]
    by_id = {s["id"]: s for s in exp["structures"]}
    ref = {"structures": {}, "esapi": {}, "warnings": []}
    for sid in sorted(set(goal_structures) | {target_id}):
        if sid not in by_id:
            ref["warnings"].append("structure %s not in export" % sid)
            continue
        d, w, outside = sample_structure(by_id[sid], z_res, smp["target_inplane_mm"], smp["z_sublevels"], dose)
        if outside:
            ref["warnings"].append("%s: %d samples outside the dose grid (dose 0)" % (sid, outside))
        ref["structures"][sid] = dvh_metrics(d, w, rx)
        ref["esapi"][sid] = dict(by_id[sid]["esapi"], volume_cm3=by_id[sid]["esapi_volume_cm3"])
    bodies = [s for s in exp["structures"] if (s["dicom_type"] or "").upper() == "EXTERNAL"]
    # The voxel scaling may come from percent of Rx; the absolute DVH maximum of the body checks it.
    if len(bodies) == 1 and bodies[0]["esapi"].get("max_gy"):
        grid_max, dvh_max = float(dose.gy.max()), bodies[0]["esapi"]["max_gy"]
        ref["dose_scale_check"] = {"voxel_unit": exp["dose"].get("voxel_unit"), "grid_max_gy": grid_max, "esapi_body_max_gy": dvh_max}
        if abs(grid_max - dvh_max) > 0.01 * dvh_max:
            ref["warnings"].append("dose grid maximum %.3f Gy differs from the absolute body DVH maximum %.3f Gy by more "
                                   "than 1 %%: dose scaling unverified" % (grid_max, dvh_max))
    tq = {}
    if len(bodies) == 1 and target_id in ref["structures"]:
        bd, bw, _ = sample_structure(bodies[0], z_res, smp["body_inplane_mm"], smp["z_sublevels"], dose,
                                     min_dose=smp["body_min_dose_fraction_of_rx"] * rx)
        v100, v50 = float((bd >= rx).sum() * bw), float((bd >= 0.5 * rx).sum() * bw)
        t = ref["structures"][target_id]
        tq = {"paddick_ci": t["v_rx_cm3"] ** 2 / (t["volume_cm3"] * v100), "gradient_index": v50 / v100,
              "body_v100_cm3": v100, "body_v50_cm3": v50}
        e, be = by_id[target_id]["esapi"], bodies[0]["esapi"]
        needed = (e.get("v_rx_cm3"), e.get("d2_gy"), e.get("d98_gy"), be.get("v_rx_cm3"), be.get("v_half_rx_cm3"))
        if None in needed or not be["v_rx_cm3"]:
            ref["warnings"].append("ESAPI direct target quality not computable from the export")
        else:
            ref["esapi_target_quality"] = {
                "paddick_ci": e["v_rx_cm3"] ** 2 / (by_id[target_id]["esapi_volume_cm3"] * be["v_rx_cm3"]),
                "gradient_index": be["v_half_rx_cm3"] / be["v_rx_cm3"],
                "homogeneity_index": (e["d2_gy"] - e["d98_gy"]) / rx}
    else:
        ref["warnings"].append("no unique EXTERNAL: CI/GI not computed")
    if target_id in ref["structures"]:
        t = ref["structures"][target_id]
        tq["homogeneity_index"] = (t["d2_gy"] - t["d98_gy"]) / rx
        tq["target_volume_cm3"] = t["volume_cm3"]
    ref["target_quality"] = tq
    if exp["treatment_orientation"] not in ("HeadFirstSupine", "HFS"):
        ref["warnings"].append("orientation %s: PAM reference not computed" % exp["treatment_orientation"])
    target = by_id.get(target_id)
    ref["beams"] = []
    for beam in exp["beams"]:
        orientation, scores = calibrate_orientation(beam, target, z_res) if target else (None, {})
        if orientation is None:
            ref["warnings"].append("beam %s: no native BEV outline of %s, PAM not computed" % (beam["id"], target_id))
        guard = tol["orientation_guard"]["max_symmetric_difference_fraction"]
        if orientation is not None and min(scores.values()) > guard:
            # Protocol deviation of 3 Nov 2026: reported, not suppressed; see PROTOKOLL.md.
            ref["warnings"].append("beam %s: best native BEV match %.2f %% exceeds the %.0f %% guard; PAM computed and "
                                   "flagged" % (beam["id"], 100 * min(scores.values()), 100 * guard))
        if orientation is not None and orientation[1]:
            ref["warnings"].append("beam %s: native BEV matched only mirrored; check axis convention" % beam["id"])
        m = beam_metrics(beam, target, z_res, mlc, orientation)
        m.update(id=beam["id"], weight_factor=beam["weight_factor"], orientation_scores=scores)
        ref["beams"].append(m)
    beams = [b for b in ref["beams"] if "error" not in b]
    total = sum(b["total_mu"] for b in beams)
    if total <= 0:
        ref["warnings"].append("no beam with MU and leaf geometry: plan aperture metrics not computed")
        ref["plan"] = {}
        return ref
    plan = {"total_mu": total, "mu_per_gy": total / (rx / exp["fractions"]),
            "mean_aperture_cm2": sum(b["mean_aperture_cm2"] * b["total_mu"] for b in beams) / total,
            "small_aperture_fraction": sum(b["small_aperture_fraction"] * b["total_mu"] for b in beams) / total}
    wf = [b for b in beams if b["weight_factor"] > 0]
    if wf and all(b["pam"] is not None for b in wf):
        plan["pam"] = sum(b["pam"] * b["weight_factor"] for b in wf) / sum(b["weight_factor"] for b in wf)
    ref["plan"] = plan
    return ref


# ---------------------------------------------------------------- comparison

GOAL_METRIC = {"D98%": "d98_gy", "Mean": "mean_gy", "Max": "max_gy", "CV20Gy": "cv20gy_cm3"}


def baseline(sqlite_path, runs_path, fractions):
    with open(runs_path, encoding="utf8") as f:
        runs = json.load(f)["runs"]
    rid = [r["review_id"] for r in runs if r["mode"] == "headless" and r["fractions"] == fractions and str(r["revision"]) == "1"]
    if len(rid) != 1:
        raise ValueError("expected exactly one headless revision-1 review for %d fractions" % fractions)
    con = sqlite3.connect(sqlite_path)
    goals = con.execute("select row_no, source_kind, resolved_structure, objective, achieved_value, unit from goal_result "
                        "where review_id=? order by row_no", rid).fetchall()
    params = con.execute("select scope, entity, name, value from plan_parameter where review_id=? order by row_no", rid).fetchall()
    return rid[0], goals, params


def within(diff, ref_value, t, rx=None):
    limits = []
    if "abs" in t:
        limits.append(t["abs"])
    if "abs_gy" in t:
        limits.append(t["abs_gy"])
    if "abs_cm3" in t:
        limits.append(t["abs_cm3"])
    if "abs_fraction_of_rx" in t:
        limits.append(t["abs_fraction_of_rx"] * rx)
    if "rel" in t:
        limits.append(t["rel"] * abs(ref_value))
    limit = max(limits)
    return abs(diff) <= limit, limit


def compare(exp, ref, goals, params, tol, gui):
    rx = exp["rx_gy"]
    ti, te = tol["independent"], tol["esapi_direct"]
    rows = []

    tg = tol["eclipse_gui"]

    def add(item, cp, ind, ind_tol, esapi=None, esapi_tol=None, gui_value=None, gui_tol=None):
        row = {"item": item, "clearplan": cp, "independent": ind, "esapi": esapi, "gui": gui_value}
        for key, value, t in (("independent", ind, ind_tol), ("esapi", esapi, esapi_tol),
                              ("gui", gui_value, gui_tol if gui_value is not None else None)):
            if t is None:
                continue
            if cp is None or value is None:
                row[key + "_verdict"] = "not comparable"
                continue
            ok, limit = within(cp - value, value, t, rx)
            row[key + "_verdict"] = "pass" if ok else "FAIL"
            row[key + "_limit"] = limit
        rows.append(row)

    for row_no, kind, structure, objective, achieved, unit in goals:
        metric = objective.split("[")[0]
        key = GOAL_METRIC.get(metric)
        if structure is None or key is None:
            add("goal %d %s %s (%s)" % (row_no, structure, objective, kind), achieved, None, None)
            continue
        ind = (ref["structures"].get(structure) or {}).get(key)
        es = ref["esapi"].get(structure, {})
        esv = {"d98_gy": es.get("d98_gy"), "mean_gy": es.get("mean_gy"), "max_gy": es.get("max_gy"),
               "cv20gy_cm3": None if es.get("v20gy_cm3") is None else es["volume_cm3"] - es["v20gy_cm3"]}[key]
        t_ind = ti["volume"] if key.endswith("cm3") else ti["dose_max"] if key == "max_gy" else ti["dose_mean_d98_d2"]
        t_es = te["volume"] if key.endswith("cm3") else te["dose"]
        add("goal %d %s %s" % (row_no, structure, objective), achieved, ind, t_ind, esv, t_es,
            gui.get((exp["label"], structure, metric)), tg["volume"] if key.endswith("cm3") else tg["dose"])

    tq, etq = ref["target_quality"], ref.get("esapi_target_quality", {})
    # Both ClearPlan and the export list treatment beams in native plan order; bind the k-th
    # pseudonymous logged beam to the k-th exported beam and accept it only if the MU agree.
    beam_refs = list(ref["beams"])
    logged = []
    for scope, entity, name, value in params:
        if scope == "beam" and name == "total_mu":
            logged.append((entity, value))
    binding = {}
    for i, (entity, mu) in enumerate(logged):
        b = beam_refs[i] if i < len(beam_refs) else None
        binding[entity] = b if b is not None and "error" not in b and abs(b["total_mu"] - mu) < 0.01 else None
    for scope, entity, name, value in params:
        if name == "estimated_duration_s" or name == "reference_dose_gy" or name == "paddick_ci_reciprocal":
            continue
        if scope == "target":
            t_ind = {"paddick_ci": ti["paddick_ci"], "gradient_index": ti["gradient_index"],
                     "homogeneity_index": ti["homogeneity_index"], "target_volume_cm3": ti["volume"]}[name]
            esv = etq.get(name) if name != "target_volume_cm3" else ref["esapi"].get(entity, {}).get("volume_cm3")
            add("%s %s" % (entity, name), value, tq.get(name), t_ind, esv,
                te["volume"] if name == "target_volume_cm3" else te["index"])
        elif scope == "plan":
            t_ind = ti[name if name in ti else "mu"]
            add("plan %s" % name, value, ref["plan"].get(name), t_ind)
        elif scope == "beam":
            b = binding.get(entity)
            add("%s %s" % (entity, name), value, None if b is None else b.get(name), ti[name if name in ti else "mu"],
                gui_value=gui.get((exp["label"], entity, name)), gui_tol=tg["mu"] if name == "total_mu" else None)
    return rows


def read_gui(path):
    if not path:
        return {}
    out = {}
    with open(path, encoding="utf8") as f:
        for r in csv.DictReader(f):
            if r.get("eclipse_value", "").strip():
                out[(r["label"], r["structure_or_beam"], r["quantity"])] = float(r["eclipse_value"].replace(",", "."))
    return out


def fmt(v):
    return "" if v is None else ("%.4g" % v if isinstance(v, float) else str(v))


def write_report(results, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    with open(os.path.join(out_dir, "k1_comparison.json"), "w", encoding="utf8") as f:
        json.dump(results, f, indent=1, default=str)
    lines = ["# K1 reference comparison", ""]
    for r in results:
        lines += ["## %s (%d fractions, ClearPlan review %s)" % (r["label"], r["fractions"], r["review_id"][:8]), ""]
        for w in r["reference"]["warnings"]:
            lines.append("- Warning: " + w)
        lines += ["", "| Item | ClearPlan | Independent | Verdict | Limit | ESAPI direct | Verdict | Eclipse GUI | Verdict |",
                  "|---|---|---|---|---|---|---|---|---|"]
        for row in r["rows"]:
            lines.append("| %s | %s | %s | %s | %s | %s | %s | %s | %s |" % (
                row["item"], fmt(row["clearplan"]), fmt(row["independent"]), row.get("independent_verdict", ""),
                fmt(row.get("independent_limit")), fmt(row["esapi"]), row.get("esapi_verdict", ""), fmt(row["gui"]),
                row.get("gui_verdict", "")))
        lines.append("")
    fails = sum(1 for r in results for row in r["rows"] for k in ("independent_verdict", "esapi_verdict", "gui_verdict") if row.get(k) == "FAIL")
    lines.append("Total failures: %d" % fails)
    with open(os.path.join(out_dir, "k1_comparison.md"), "w", encoding="utf8") as f:
        f.write("\n".join(lines) + "\n")
    return fails


def main(argv=None):
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("compare")
    c.add_argument("--export", action="append", required=True)
    c.add_argument("--label-fractions", nargs="+", required=True)
    c.add_argument("--sqlite", required=True)
    c.add_argument("--runs", required=True)
    c.add_argument("--tolerances", required=True)
    c.add_argument("--mlc", required=True)
    c.add_argument("--gui")
    c.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    tol = json.load(open(a.tolerances, encoding="utf8"))
    mlc = json.load(open(a.mlc, encoding="utf8"))
    fractions = {k: int(v) for k, v in (x.split("=") for x in a.label_fractions)}
    gui = read_gui(a.gui)
    results = []
    for path in a.export:
        exp = load_export(path)
        fx = fractions[exp["label"]]
        if exp["fractions"] != fx:
            raise ValueError("export %s has %d fractions, expected %d" % (path, exp["fractions"], fx))
        rid, goals, params = baseline(a.sqlite, a.runs, fx)
        target_id = next(e for s, e, n, v in params if s == "target" and n == "target_volume_cm3")
        goal_structures = [g[2] for g in goals if g[2]]
        ref = reference(exp, tol, mlc, target_id, goal_structures)
        results.append({"label": exp["label"], "fractions": fx, "review_id": rid, "export": os.path.basename(path),
                        "context_sha256": exp["context_sha256"], "reference": ref,
                        "rows": compare(exp, ref, goals, params, tol, gui)})
    fails = write_report(results, a.out)
    print("failures:", fails)
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
