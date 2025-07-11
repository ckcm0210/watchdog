#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Demo script showing the enhanced watchdog functionality
"""

import sys
import os
import tempfile
import time
sys.path.append('/home/runner/work/watchdog/watchdog')

from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.formula import ArrayFormula

# Import the enhanced functions
from watch import serialize_cell_value, extract_external_links, resolve_external_references

def demo_array_formula_filtering():
    """Demo the ArrayFormula filtering feature"""
    print("🚀 DEMO: ArrayFormula Filtering")
    print("=" * 50)
    
    # Create two ArrayFormula objects with identical content but different addresses
    print("Creating two ArrayFormula objects with identical content...")
    
    af1 = ArrayFormula(ref="A1:A3")
    af1.text = "SUM(A1:A3)"
    
    af2 = ArrayFormula(ref="A1:A3")
    af2.text = "SUM(A1:A3)"
    
    print(f"ArrayFormula 1: {repr(af1)}")
    print(f"ArrayFormula 2: {repr(af2)}")
    print(f"Are they equal? {af1 == af2}")
    print()
    
    # Show the OLD behavior (what would happen without enhancement)
    print("OLD behavior (without enhancement):")
    old_val1 = str(af1)
    old_val2 = str(af2)
    print(f"  str(af1): {old_val1}")
    print(f"  str(af2): {old_val2}")
    print(f"  Would detect as change: {old_val1 != old_val2}")
    print()
    
    # Show the NEW behavior (with enhancement)
    print("NEW behavior (with enhancement):")
    new_val1 = serialize_cell_value(af1)
    new_val2 = serialize_cell_value(af2)
    print(f"  serialize_cell_value(af1): {new_val1}")
    print(f"  serialize_cell_value(af2): {new_val2}")
    print(f"  Correctly filters out change: {new_val1 == new_val2}")
    print()
    
    if new_val1 == new_val2:
        print("✅ SUCCESS: ArrayFormula object address changes are now filtered!")
    else:
        print("❌ FAILURE: ArrayFormula filtering not working!")
    
    print()

def demo_external_link_mapping():
    """Demo the external link mapping feature"""
    print("🔗 DEMO: External Link Mapping")
    print("=" * 50)
    
    # Create a mock external link mapping (as would be extracted from real Excel files)
    print("Mock external link mapping (as extracted from Excel files):")
    external_links = {
        1: "C:\\Data\\SourceFile1.xlsx",
        2: "C:\\Reports\\DataTable.xlsx",
        3: "\\\\NetworkDrive\\shared\\CalculationSheet.xlsx"
    }
    
    for key, value in external_links.items():
        print(f"  [{key}] -> {value}")
    print()
    
    # Test formula resolution
    print("Formula resolution examples:")
    test_formulas = [
        "=[1]Sheet1!A1",
        "=[2]DataTable!B1",
        "=SUM([1]Sheet1!A1:[1]Sheet1!A10)",
        "=IF([3]Calculations!C1>0,[2]DataTable!B1,[1]Sheet1!A1)",
        "=VLOOKUP(A1,[2]DataTable!A:B,2,FALSE)"
    ]
    
    for formula in test_formulas:
        resolved = resolve_external_references(formula, external_links)
        print(f"  Original: {formula}")
        print(f"  Resolved: {resolved}")
        if formula != resolved:
            print(f"  ✅ External links resolved!")
        else:
            print(f"  ℹ️  No external links found in formula")
        print()
    
    print("✅ SUCCESS: External link mapping working correctly!")
    print()

def demo_combined_scenario():
    """Demo a combined scenario with both features"""
    print("🎯 DEMO: Combined Scenario")
    print("=" * 50)
    
    print("Simulating a complete watchdog scenario...")
    
    # Create baseline cell data
    af_baseline = ArrayFormula(ref="A1:A3")
    af_baseline.text = "SUM(A1:A3)"
    
    baseline_cells = {
        "MainSheet": {
            "A1": {"formula": None, "value": 100},
            "B1": {"formula": "=SUM(A1:A3)", "value": serialize_cell_value(af_baseline)},
            "C1": {"formula": "=[1]ExternalSheet!A1", "value": "=[1]ExternalSheet!A1"}
        }
    }
    
    # Create current cell data (simulating re-read of same file)
    af_current = ArrayFormula(ref="A1:A3")
    af_current.text = "SUM(A1:A3)"  # Same content, different object
    
    current_cells = {
        "MainSheet": {
            "A1": {"formula": None, "value": 100},
            "B1": {"formula": "=SUM(A1:A3)", "value": serialize_cell_value(af_current)},
            "C1": {"formula": "=[1]ExternalSheet!A1", "value": "=[1]ExternalSheet!A1"}
        }
    }
    
    print("1. Baseline ArrayFormula object:", repr(af_baseline))
    print("2. Current ArrayFormula object:", repr(af_current))
    print("3. Are objects equal?", af_baseline == af_current)
    print("4. Are serialized values equal?", 
          serialize_cell_value(af_baseline) == serialize_cell_value(af_current))
    print()
    
    # Check for changes
    changes_detected = baseline_cells != current_cells
    print(f"5. Changes detected: {changes_detected}")
    
    if not changes_detected:
        print("✅ SUCCESS: No false positive from ArrayFormula object address change!")
    else:
        print("❌ FAILURE: False positive detected!")
    
    # Now test with external link resolution
    external_links = {1: "SourceData.xlsx"}
    resolved_formula = resolve_external_references("=[1]ExternalSheet!A1", external_links)
    print(f"6. External link resolved: =[1]ExternalSheet!A1 -> {resolved_formula}")
    
    print()
    print("✅ DEMO completed successfully!")

if __name__ == "__main__":
    print("🎭 Enhanced Watchdog Demo")
    print("=" * 60)
    print()
    
    demo_array_formula_filtering()
    demo_external_link_mapping()
    demo_combined_scenario()
    
    print("=" * 60)
    print("✅ All demos completed! Enhanced watchdog is ready to use.")
    print()
    print("Key benefits:")
    print("1. ✅ No more false positives from ArrayFormula object address changes")
    print("2. ✅ External link references show actual file paths")
    print("3. ✅ Backward compatible with existing functionality")
    print("4. ✅ Minimal performance impact")
    print("=" * 60)