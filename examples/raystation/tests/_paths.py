"""Import paths for the tests (no installation needed)."""
import os
import sys

TESTS = os.path.dirname(os.path.abspath(__file__))
RAYSTATION = os.path.dirname(TESTS)
SRC = os.path.dirname(os.path.dirname(RAYSTATION))
REVIEW_LOG_TOOLS = os.path.join(SRC, "tools", "review-log")
RULES = os.path.join(RAYSTATION, "rules")
for path in (RAYSTATION, TESTS, REVIEW_LOG_TOOLS):
    if path not in sys.path:
        sys.path.insert(0, path)
