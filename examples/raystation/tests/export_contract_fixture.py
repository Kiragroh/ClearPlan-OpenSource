"""Export one deterministic fake-context RayStation snapshot for contract tests."""

import sys
from datetime import datetime, timezone
from pathlib import Path

from test_clearplan_review_snapshot import ADAPTER, build_context


def main() -> int:
    if len(sys.argv) != 2:
        raise SystemExit("Usage: export_contract_fixture.py OUTPUT_JSON")

    output = Path(sys.argv[1]).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    _, _, get_current, _ = build_context()
    ADAPTER.export_current_review_snapshot(
        output_path=output,
        get_current_fn=get_current,
        generated_utc=datetime(2026, 7, 30, 8, 0, tzinfo=timezone.utc),
        overwrite=True,
    )
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
