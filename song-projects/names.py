# -*- coding: utf-8 -*-
"""אילו פרויקטים נמכרו (בלי שמות; השמות שמורים בקובץ פרטי מחוץ לריפו)."""
import os, sys

TAKEN_IDS = ["01", "02", "03", "04", "07", "08", "09"]
NAMES, SINGULAR = {}, {}
if os.environ.get("NAMES_DIR"):
    sys.path.insert(0, os.environ["NAMES_DIR"])
    from names_private import NAMES, SINGULAR  # noqa: F401
