"""Assemble reviewed UTF-8 manuscript sources without editing the originals."""
import argparse
import json
import re
from pathlib import Path

FIGURE_2_CAPTION = (
    "**Figure 2. Synthetic review of estimated delivery parameters and field geometry.** "
    "Actual application captures from the `publication-dual-layer` fixture in [FINAL_RELEASE] "
    "show (A) the nominal field setting and estimated MU rate against control-point index, "
    "with the displayed 12 degrees/s gantry-speed assumption, and (B) the first-control-point "
    "beam's-eye view with two staggered, jawless multileaf-collimator layers, isocenter-plane "
    "coordinates, a synthetic CT proxy, and corresponding field metadata. The rate curve is "
    "an estimate, not a measured delivery trace. All anatomy and values are synthetic; the "
    "aperture geometry does not generate the analytical dose. The GUI labels are retained "
    "as captured. These views illustrate the available review context, not clinical "
    "performance or measured usability."
)

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--body", type=Path, required=True)
parser.add_argument("--abstract", type=Path, required=True)
parser.add_argument("--output", type=Path, required=True)
parser.add_argument("--figure-directory", type=Path, required=True)
parser.add_argument("--release-metadata", type=Path)
args = parser.parse_args()
body = args.body.read_text(encoding="utf-8")
abstract = args.abstract.read_text(encoding="utf-8").split("## Quality note", 1)[0].strip()
abstract = re.sub(r"^# Abstract", "## Abstract", abstract)
body = body.replace("# ClearPlan: configurable", "# ClearPlan configurable", 1)
body = re.sub(r"Article type:.*", "Article type: Technical note", body, count=1)
body = re.sub(r"<!-- Revision-phase candidate\..*?-->", abstract, body, count=1, flags=re.DOTALL)
body = body.replace(
    "Subsequent readback checks patient, provider, type, status, and returned attachment bytes or fingerprint, distinguishing metadata-only from full attachment verification.",
    "Subsequent readback checks patient, provider, type, status, document date, and returned attachment bytes or fingerprint. Where ARIA returns only a file reference, an optional reader is restricted to explicitly configured UNC roots and verifies the PDF bytes without modifying the stored file. Missing attachment creation metadata is disclosed. Metadata-only and full attachment verification remain distinct.")
body = body.replace(
    "The authors reviewed and edited all generated content and take full responsibility for the published work.",
    "The authors take responsibility for the manuscript and associated software documentation.")
body = body.replace(" [Retain the completed author-review statement only after the authors have performed that review.]", "")
body = body.replace("Final source commit: [FINAL_COMMIT].",
                    "Verified software source commit: [FINAL_COMMIT]. The versioned release additionally contains the manuscript and publication artifacts generated from that software revision.")
body, caption_count = re.subn(r"(?ms)^\*\*Figure 2\..*?(?=\n\s*\n|\Z)",
                              lambda _: FIGURE_2_CAPTION, body, count=1)
if caption_count != 1:
    raise ValueError("Expected one Figure 2 caption in the reviewed manuscript source.")
for number, stem in [(1, "Figure_1_ClearPlan_architecture"), (2, "Figure_2_ClearPlan_workspace")]:
    image = (args.figure_directory / (stem + ".png")).resolve()
    if not image.is_file():
        raise FileNotFoundError(image)
    import os
    relative = Path(os.path.relpath(image, args.output.resolve().parent)).as_posix()
    anchor = "**Figure " + str(number) + "."
    body = body.replace(anchor, f"![Figure {number}]({relative})\n\n" + anchor, 1)
if args.release_metadata:
    metadata = json.loads(args.release_metadata.read_text(encoding="utf-8"))
    required = {"FINAL_RELEASE", "FINAL_COMMIT", "FINAL_RELEASE_URL", "CORE_TESTS", "FINAL_TEST_RESULT", "SUPPLEMENT_HASH",
                "FINAL_PUBLICATION_EXPORT_VERIFICATION", "FINAL_OUTPUT_DESCRIPTION"}
    if not required <= metadata.keys():
        raise ValueError("Release metadata is incomplete.")
    if not re.fullmatch(r"[0-9a-f]{40}", metadata["FINAL_COMMIT"]) or not re.fullmatch(r"[0-9a-f]{64}", metadata["SUPPLEMENT_HASH"]):
        raise ValueError("Source commit and supplement hash must be exact.")
    for key in required:
        body = re.sub(r"\[" + key + r"(?::[^\]]*)?\]", lambda _: str(metadata[key]), body)
    body = body.replace(" These placeholders must be replaced by published and verified artifacts before submission.", "")
    body = body.replace("<!--anchor:none:-->", "<!--anchor:release:" + metadata["FINAL_RELEASE"] + "-->")
args.output.parent.mkdir(parents=True, exist_ok=True)
args.output.write_text(body, encoding="utf-8")
print("Assembled manuscript; unresolved final markers:", len(re.findall(r"\[(?:FINAL_|CORE_TESTS|SUPPLEMENT_HASH)", body)))
