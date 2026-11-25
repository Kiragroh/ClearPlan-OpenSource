"""Synthetic checks of k1_reference.py with analytically known answers.

Run: python -I test_k1_reference.py
"""
import json
import math
import os
import sys
import tempfile
import unittest

import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import k1_reference as k  # noqa: E402
import k1_mlc_check as mc  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
TOL = json.load(open(os.path.join(HERE, "..", "tolerances.json"), encoding="utf8"))
MLC = json.load(open(os.path.join(HERE, "..", "mlc_geometry.json"), encoding="utf8"))
Z_RES = 2.5


def sphere_contours(radius, centre=(0, 0, 0), n=180):
    out = []
    zs = np.arange(-40, 40.01, Z_RES)
    for plane, z in enumerate(zs):
        dz = z - centre[2]
        if abs(dz) >= radius:
            continue
        r = math.sqrt(radius ** 2 - dz ** 2)
        t = np.linspace(0, 2 * np.pi, n, endpoint=False)
        out.append({"plane": plane, "xyz": [[centre[0] + r * math.cos(a), centre[1] + r * math.sin(a), z] for a in t]})
    return out


def box_contours(half):
    zs = np.arange(-40, 40.01, Z_RES)
    return [{"plane": p, "xyz": [[-half, -half, z], [half, -half, z], [half, half, z], [-half, half, z]]}
            for p, z in enumerate(zs)]


def radial_dose(r):
    # Strictly decreasing cone: 50 Gy at the centre, Rx (25 Gy) at 25 mm, half Rx at 37.5 mm.
    return np.clip(50 * (1 - r / 50), 0, None)


def make_export(tmp, beams, target_centre=(0, 0, 0)):
    res = 1.0
    n = 101
    origin = [-50, -50, -50]
    ax = np.arange(n) * res - 50
    z, y, x = np.meshgrid(ax, ax, ax, indexing="ij")
    gy = radial_dose(np.sqrt(x ** 2 + y ** 2 + z ** 2))
    raw = np.round(gy * 1e5).astype("<i4")
    raw.tofile(os.path.join(tmp, "syn.dose.bin"))
    exp = {
        "label": "single", "context_sha256": "0", "rx_gy": 25.0, "fractions": 5, "treatment_orientation": "HeadFirstSupine",
        "dose": {"file": "syn.dose.bin", "gy_per_raw": 1e-5, "gy_offset": 0.0, "size": [n, n, n], "res_mm": [res] * 3,
                 "origin": origin, "x_dir": [1, 0, 0], "y_dir": [0, 1, 0], "z_dir": [0, 0, 1]},
        "image": {"z_res_mm": Z_RES},
        "structures": [
            {"id": "PTV", "dicom_type": "PTV", "esapi_volume_cm3": 0, "esapi": {}, "contours": sphere_contours(20, target_centre)},
            {"id": "BODY", "dicom_type": "EXTERNAL", "esapi_volume_cm3": 0, "esapi": {}, "contours": box_contours(48)},
        ],
        "beams": beams,
    }
    path = os.path.join(tmp, "syn.json")
    with open(path, "w", encoding="utf8") as f:
        json.dump(exp, f)
    return path


def hd120_cp(gantry, collimator, x1, x2, weight, jaws=(-100, -30, 100, 30)):
    sad = 1000.0
    a = math.radians(gantry)
    return {"meterset_weight": weight, "gantry": gantry, "collimator": collimator, "couch": 0.0,
            "jaws_x1_y1_x2_y2": list(jaws), "source": [sad * math.sin(a), -sad * math.cos(a), 0.0],
            "leaves": [[x1] * 60, [x2] * 60]}


def beam(cps, mlc="Varian High Definition 120", outlines=None):
    return {"id": "B", "mlc_model": mlc, "meterset_mu": 100.0, "weight_factor": 1.0, "isocenter": [0, 0, 0],
            "control_points": cps, "native_bev_outlines": outlines or []}


class Synthetic(unittest.TestCase):
    def test_dvh_and_target_quality(self):
        with tempfile.TemporaryDirectory() as tmp:
            exp = k.load_export(make_export(tmp, []))
            ref = k.reference(exp, TOL, MLC, "PTV", ["PTV"])
            t = ref["structures"]["PTV"]
            v_target = 4 / 3 * math.pi * 20 ** 3 / 1000
            self.assertAlmostEqual(t["volume_cm3"], v_target, delta=0.02 * v_target)
            # Dose 50 - r over a sphere of radius 20: mean 50 - 3/4*20; D_v at r = 20*(1-v)^(1/3).
            self.assertAlmostEqual(t["mean_gy"], 35, delta=0.1)
            # Staircase slabs of 2.5 mm shift the edge by a fraction of a millimetre (1 Gy/mm here);
            # the bound is the pre-registered 1 % of Rx.
            self.assertAlmostEqual(t["d98_gy"], 50 - 20 * 0.98 ** (1 / 3), delta=0.25)
            self.assertAlmostEqual(t["d2_gy"], 50 - 20 * 0.02 ** (1 / 3), delta=0.25)
            tq = ref["target_quality"]
            v100, v50 = 4 / 3 * math.pi * 25 ** 3 / 1000, 4 / 3 * math.pi * 37.5 ** 3 / 1000
            self.assertAlmostEqual(tq["body_v100_cm3"], v100, delta=0.02 * v100)
            self.assertAlmostEqual(tq["gradient_index"], v50 / v100, delta=0.02 * v50 / v100)
            self.assertAlmostEqual(tq["paddick_ci"], v_target / v100, delta=0.03 * v_target / v100)
            hi = (20 * 0.98 ** (1 / 3) - 20 * 0.02 ** (1 / 3)) / 25
            self.assertAlmostEqual(tq["homogeneity_index"], hi, delta=0.01)

    def test_aperture_and_pam_open_and_half_blocked(self):
        open_beam = beam([hd120_cp(0, 0, -40, 40, 0), hd120_cp(0, 0, -40, 40, 1)])
        half_beam = beam([hd120_cp(0, 0, -40, 0, 0), hd120_cp(0, 0, -40, 0, 1)])
        with tempfile.TemporaryDirectory() as tmp:
            exp = k.load_export(make_export(tmp, []))
            target = exp["structures"][0]
            m = k.beam_metrics(open_beam, target, Z_RES, MLC, (1, False))
            self.assertAlmostEqual(m["mean_aperture_cm2"], 80 * 60 / 100, places=6)
            self.assertAlmostEqual(m["pam"], 0, delta=1e-9)
            m = k.beam_metrics(half_beam, target, Z_RES, MLC, (1, False))
            self.assertAlmostEqual(m["pam"], 0.5, delta=0.01)
            self.assertAlmostEqual(m["small_aperture_fraction"], 0)

    def test_sx2_layers_intersect(self):
        cp = hd120_cp(0, 0, 0, 0, 1)
        leaves = np.zeros((2, 57))
        leaves[0, :], leaves[1, :] = -30, 30
        leaves[:, 0:28] = 0
        leaves[0, 12:16], leaves[1, 12:16] = -30, 30  # distal y -20..20 open
        cp["leaves"] = leaves.tolist()
        ap = k.aperture(cp, MLC["models"]["SX2"])
        self.assertAlmostEqual(ap.area, 60 * 40, places=6)

    def test_orientation_check_detects_mirroring(self):
        # The native outline is unrotated by the collimator, so only mirroring is decidable.
        centre = (15, 0, 8)
        cps = [hd120_cp(60, 30, -40, 40, 0), hd120_cp(60, 30, -40, 40, 1)]
        with tempfile.TemporaryDirectory() as tmp:
            exp = k.load_export(make_export(tmp, [], centre))
            target = exp["structures"][0]
            for mirror in (False, True):
                sil = k.silhouette(target, Z_RES, cps[0]["source"], [0, 0, 0], 0, 1, mirror)
                native = [{"structure": "PTV", "polygons": [np.asarray(sil.exterior.coords).ravel().tolist()]}]
                best, scores = k.calibrate_orientation(beam(cps, outlines=native), target, Z_RES)
                self.assertEqual(best, (k.COLLIMATOR_SIGN, mirror))
                self.assertLess(min(scores.values()), 1e-6)
                self.assertGreater(max(scores.values()), 0.05)

    def test_collimator_rotation_is_counter_clockwise_from_source(self):
        # Gantry 0, source above: positive collimator angle turns the collimator X axis towards +Y (superior).
        d, sad, xc, yc = k.bev_axes([0, -1000, 0], [0, 0, 0], 90, k.COLLIMATOR_SIGN)
        np.testing.assert_allclose(xc, [0, 0, 1], atol=1e-12)
        np.testing.assert_allclose(yc, [-1, 0, 0], atol=1e-12)

    def test_mlc_check_finds_geometry_from_asymmetric_jaws(self):
        # Pairs open exactly where their HD120 band overlaps Y jaws -24..36 mm (leaves 20-44).
        cp = hd120_cp(0, 0, 0, 0, 1, jaws=(-30, -24, 30, 36))
        cp["leaves"][0][20:45] = [-10] * 25
        cp["leaves"][1][20:45] = [10] * 25
        scores, best = mc.check_hd120(beam([cp]), MLC["models"]["Varian High Definition 120"])
        self.assertEqual(best, "spec (mlc_geometry.json)")
        self.assertEqual(scores[best]["open_pairs_outside_jaws"], 0)
        self.assertGreater(scores["index order reversed"]["open_pairs_outside_jaws"], 0)
        self.assertGreater(scores["shifted +2.5 mm"]["open_pairs_outside_jaws"], 0)
        self.assertGreater(scores["shifted -2.5 mm"]["open_pairs_outside_jaws"], 0)

    def test_divergent_silhouette_magnifies(self):
        with tempfile.TemporaryDirectory() as tmp:
            exp = k.load_export(make_export(tmp, []))
            target = exp["structures"][0]
            sil = k.silhouette(target, Z_RES, [0, -1000, 0], [0, 0, 0], 0, 1)
            # Sphere r=20 at the isocentre: silhouette radius r*SAD/sqrt(SAD^2-r^2); slab ends add < 1.3 mm in y.
            self.assertGreater(sil.area, math.pi * 20 ** 2 * 0.99)
            self.assertLess(sil.area, math.pi * 20 ** 2 * 1.10)


if __name__ == "__main__":
    unittest.main(verbosity=2)
