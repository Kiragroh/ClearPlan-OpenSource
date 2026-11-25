"""Plan-data check of mlc_geometry.json against the native K1 exports.

ESAPI 18 exposes leaf positions but no leaf boundaries (no leaf-width or boundary member
in VMS.TPS.Common.Model.API/Types), so the geometry cannot be read from Eclipse. Instead,
candidate geometries are tested against what Eclipse did with the leaves:

- HD120 (physical jaws): Eclipse closes every leaf pair whose band lies wholly outside
  the Y jaws. An open pair (gap > 1 mm) outside the jaws under a candidate geometry is a
  contradiction. The phantom's Y jaws are asymmetric, so this also fixes index direction.
- SX2 (jawless): the open region of each layer should be centred on the target's
  beam's-eye-view silhouette (MU-weighted offset of each layer's open-band centre), and
  the effective aperture should lie over the target (MU-weighted share of open area more
  than 5 mm outside the silhouette). A candidate must be centred within half a leaf in
  both layers; among those, the least open area beside the target wins.

Usage:
  python k1_mlc_check.py --export <k1-*.json> [--export ...] --target-single PTV
      --target-double PTV_High --mlc mlc_geometry.json --out <dir>
"""
import argparse
import json
import os
import sys

import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import k1_reference as k  # noqa: E402

OPEN_GAP_MM = 1.0  # closed pairs sit at 0.5 mm (HD120) or 0 mm (SX2)
JAW_TOL_MM = 0.1
AREA_MARGIN_MM = 5.0


def hd120_hypotheses(spec):
    b = np.asarray(spec["layers"][0]["boundaries_mm"], dtype=float)
    millennium = np.concatenate([np.arange(-200, -100, 10), np.arange(-100, 100, 5), np.arange(100, 200.1, 10)])
    # (boundaries, leaf order flipped)
    return {"spec (mlc_geometry.json)": (b, False), "index order reversed": (b, True),
            "shifted +2.5 mm": (b + 2.5, False), "shifted -2.5 mm": (b - 2.5, False),
            "Millennium 120 widths": (millennium, False)}


def sx2_hypotheses(spec):
    dist, prox = (np.asarray(l["boundaries_mm"], dtype=float) for l in spec["layers"])
    # (distal leaf range, distal boundaries, proximal leaf range, proximal boundaries, leaf order flipped)
    return {
        "spec (mlc_geometry.json)": ((0, 28), dist, (28, 57), prox, False),
        "layers swapped (proximal 0-28)": ((29, 57), dist, (0, 29), prox, False),
        "index order reversed": ((0, 28), dist, (28, 57), prox, True),
        "distal shifted +5 mm": ((0, 28), dist + 5, (28, 57), prox, False),
        "proximal shifted +5 mm": ((0, 28), dist, (28, 57), prox + 5, False),
    }


def gaps(cp):
    leaves = np.asarray(cp["leaves"], dtype=float)
    return leaves[1] - leaves[0]


def check_hd120(beam, spec):
    cps = beam["control_points"]
    out = {}
    for name, (b, flip) in hd120_hypotheses(spec).items():
        violations, cps_hit = 0, 0
        for cp in cps:
            _, y1, _, y2 = cp["jaws_x1_y1_x2_y2"]
            g = gaps(cp)
            is_open = (g[::-1] if flip else g) > OPEN_GAP_MM
            outside = (b[1:] < y1 - JAW_TOL_MM) | (b[:-1] > y2 + JAW_TOL_MM)
            n = int((is_open & outside).sum())
            violations += n
            cps_hit += n > 0
        out[name] = {"open_pairs_outside_jaws": violations, "control_points_with_contradiction": cps_hit,
                     "control_points": len(cps)}
    return out, min(out, key=lambda n: out[n]["open_pairs_outside_jaws"])


def check_sx2(beam, spec, target, z_res, orientation):
    w = k.mu_weights(beam)
    hyp = sx2_hypotheses(spec)
    # weighted distal offset, proximal offset, open area outside target + margin, total open area, target covered, weight
    acc = {name: np.zeros(6) for name in hyp}
    for cp, wi in zip(beam["control_points"], w):
        if wi <= 0:
            continue
        g = gaps(cp)
        leaves = np.asarray(cp["leaves"], dtype=float)
        sil = k.silhouette(target, z_res, cp["source"], beam["isocenter"], cp["collimator"], *orientation)
        near = sil.buffer(AREA_MARGIN_MM)
        centre = (sil.bounds[1] + sil.bounds[3]) / 2
        for name, (d_idx, d_b, p_idx, p_b, flip) in hyp.items():
            offs = []
            for (a, e), bounds in ((d_idx, d_b), (p_idx, p_b)):
                layer = g[a:e][::-1] if flip else g[a:e]
                open_local = np.where(layer > OPEN_GAP_MM)[0]
                if open_local.size == 0:
                    break
                offs.append((bounds[open_local.min()] + bounds[open_local.max() + 1]) / 2 - centre)
            if len(offs) < 2:
                continue
            lv = leaves.copy()
            if flip:
                for a, e in (d_idx, p_idx):
                    lv[:, a:e] = leaves[:, a:e][:, ::-1]
            model = {"jaws": "none", "layers": [{"leaf_indices": list(d_idx), "boundaries_mm": d_b.tolist()},
                                                {"leaf_indices": list(p_idx), "boundaries_mm": p_b.tolist()}]}
            ap = k.aperture(dict(cp, leaves=lv.tolist()), model)
            acc[name] += wi * np.array([offs[0], offs[1], ap.difference(near).area, ap.area,
                                        ap.intersection(sil).area / sil.area, 1])
    out = {name: {"distal_open_centre_offset_mm": a[0] / a[5], "proximal_open_centre_offset_mm": a[1] / a[5],
                  "open_area_outside_target_plus_margin": a[2] / a[3], "target_covered": a[4] / a[5]}
           for name, a in acc.items()}
    # Centred within half a leaf in both layers, then the least open area beside the target.
    centred = [n for n, v in out.items() if max(abs(v["distal_open_centre_offset_mm"]),
                                                abs(v["proximal_open_centre_offset_mm"])) < 5.0]
    best = min(centred, key=lambda n: out[n]["open_area_outside_target_plus_margin"]) if centred else None
    return out, best

def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--export", action="append", required=True)
    ap.add_argument("--target-single", default="PTV")
    ap.add_argument("--target-double", default="PTV_High")
    ap.add_argument("--mlc", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    mlc = json.load(open(a.mlc, encoding="utf8"))
    results = {}
    for path in a.export:
        exp = k.load_export(path)
        target_id = a.target_single if exp["label"] == "single" else a.target_double
        target = next(s for s in exp["structures"] if s["id"] == target_id)
        z_res = exp["image"]["z_res_mm"]
        for i, beam in enumerate(exp["beams"]):
            model = beam["mlc_model"]
            spec = mlc["models"][model]
            if spec["jaws"] == "physical":
                scores, best = check_hd120(beam, spec)
            else:
                orientation, _ = k.calibrate_orientation(beam, target, z_res)
                scores, best = check_sx2(beam, spec, target, z_res, orientation or (k.COLLIMATOR_SIGN, False))
            results["%s beam %d" % (exp["label"], i + 1)] = {
                "mlc_model_from_esapi": model, "leaf_pairs_in_export": len(beam["control_points"][0]["leaves"][0]),
                "leaf_pairs_in_spec": spec["layers"][-1]["leaf_indices"][1], "scores": scores, "best": best}
    os.makedirs(a.out, exist_ok=True)
    with open(os.path.join(a.out, "k1_mlc_check.json"), "w", encoding="utf8") as f:
        json.dump(results, f, indent=1)
    lines = ["# K1 MLC geometry check from native leaf positions", ""]
    for key, r in results.items():
        lines += ["## " + key, "", "ESAPI MLC model: %s; leaf pairs export/spec: %d/%d; best: %s" % (
            r["mlc_model_from_esapi"], r["leaf_pairs_in_export"], r["leaf_pairs_in_spec"], r["best"]), ""]
        first = next(iter(r["scores"].values()))
        if "open_pairs_outside_jaws" in first:
            lines += ["| Geometry | Open pairs outside Y jaws | Control points affected |", "|---|---:|---:|"]
            lines += ["| %s | %d | %d of %d |" % (n, v["open_pairs_outside_jaws"], v["control_points_with_contradiction"],
                                                 v["control_points"]) for n, v in r["scores"].items()]
        else:
            lines += ["| Geometry | Distal centre offset (mm) | Proximal centre offset (mm) | Open area > 5 mm outside target | Target covered |",
                      "|---|---:|---:|---:|---:|"]
            lines += ["| %s | %+.1f | %+.1f | %.3f | %.3f |" % (
                n, v["distal_open_centre_offset_mm"], v["proximal_open_centre_offset_mm"],
                v["open_area_outside_target_plus_margin"], v["target_covered"]) for n, v in r["scores"].items()]
        lines.append("")
    with open(os.path.join(a.out, "k1_mlc_check.md"), "w", encoding="utf8") as f:
        f.write("\n".join(lines))
    print("\n".join(lines))


if __name__ == "__main__":
    main()
