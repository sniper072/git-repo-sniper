#!/usr/bin/env python3
"""Compare compact vs field layout on two device profiles."""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from template_engine import compare_layouts_on_devices

TEST_DEVICES = [
    pd.Series(
        {
            "inventory": "67+",
            "model": "Kyocera ECOSYS M3040dn",
            "serial": "LS67Y43279",
            "location": "Риелторы (Тюмень)",
            "office": "Тюмень",
            "unit": "Риелторы",
            "employee": "",
            "prev_counter": "272339",
            "report_counter": "272864",
            "print_count": "525",
            "row_number": "1",
        }
    ),
    pd.Series(
        {
            "inventory": "HP-204",
            "model": "HP Color LaserJet Pro M454dn",
            "serial": "PHC8K99102",
            "location": "Колл-центр (Екатеринбург)",
            "office": "Екатеринбург",
            "unit": "Колл-центр",
            "employee": "Сидоров П.А.",
            "prev_counter": "118450",
            "report_counter": "119102",
            "print_count": "652",
            "row_number": "42",
        }
    ),
]


def main() -> None:
    output_dir = Path("equipment-templates/output/download/compare")
    template_dir = Path("equipment-templates/input")
    winner, totals = compare_layouts_on_devices(TEST_DEVICES, template_dir, output_dir)
    print(f"Winner: {winner}")
    print(f"Totals: compact={totals['compact']} field={totals['field']}")
    print(f"Report: {output_dir / 'LAYOUT_COMPARE_REPORT.txt'}")


if __name__ == "__main__":
    main()
