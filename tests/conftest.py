"""Make the pipeline modules importable from the tests.

The stages live in separate folders and are run as scripts, so they are not
an installed package. Adding their folders to the path is enough for pytest.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

for folder in ("shared",
               "01-supplier-intake/src",
               "02-requisition-builder/src",
               "04-po-handover/src"):
    sys.path.insert(0, str(ROOT / folder))
