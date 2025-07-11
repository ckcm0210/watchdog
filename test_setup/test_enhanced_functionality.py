#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test the enhanced watchdog functionality
"""

import sys
import os
import tempfile
import shutil
sys.path.append('/home/runner/work/watchdog/watchdog')

from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.formula import ArrayFormula

# Import the enhanced functions
from watch import serialize_cell_value, extract_external_links, resolve_external_references, dump_excel_cells_with_timeout

def test_array_formula_filtering():
    """Test ArrayFormula filtering functionality"""
    print("=" * 60)
    print("Testing ArrayFormula filtering")
    print("=" * 60)
    
    # Create two ArrayFormula objects with same content
    af1 = ArrayFormula(ref="A1:A3")
    af1.text = "SUM(A1:A3)"
    
    af2 = ArrayFormula(ref="A1:A3")
    af2.text = "SUM(A1:A3)"
    
    print(f"Original object representation:")
    print(f"  af1: {repr(af1)}")
    print(f"  af2: {repr(af2)}")
    print(f"  Objects equal: {af1 == af2}")
    print(f"  Object addresses equal: {repr(af1) == repr(af2)}")
    
    print(f"\nSerialized values:")
    val1 = serialize_cell_value(af1)
    val2 = serialize_cell_value(af2)
    print(f"  af1 serialized: {val1}")
    print(f"  af2 serialized: {val2}")
    print(f"  Serialized values equal: {val1 == val2}")
    
    # Test with different formulas
    af3 = ArrayFormula(ref="B1:B3")
    af3.text = "SUM(B1:B3)"
    
    val3 = serialize_cell_value(af3)
    print(f"\nDifferent formula test:")
    print(f"  af3 serialized: {val3}")
    print(f"  af1 == af3: {val1 == val3}")
    
    print("\n✅ ArrayFormula filtering test completed")

def test_external_link_mapping():
    """Test external link mapping functionality"""
    print("\n" + "=" * 60)
    print("Testing external link mapping")
    print("=" * 60)
    
    # Test with our test file
    test_file = '/home/runner/work/watchdog/watchdog/test_setup/test_external_links.xlsx'
    
    if os.path.exists(test_file):
        print(f"Testing with file: {test_file}")
        
        # Test external link extraction
        external_links = extract_external_links(test_file)
        print(f"External links found: {external_links}")
        
        # Test formula resolution
        test_formulas = [
            "=[1]Sheet1!A1",
            "=[2]Data!B1",
            "=[3]Table!C1",
            "=SUM([1]Sheet1!A1:[1]Sheet1!A10)",
            "=IF([1]Sheet1!A1>0,[2]Data!B1,0)"
        ]
        
        print(f"\nFormula resolution test:")
        for formula in test_formulas:
            resolved = resolve_external_references(formula, external_links)
            print(f"  Original: {formula}")
            print(f"  Resolved: {resolved}")
            if formula != resolved:
                print(f"  -> 已解析外部連結！")
            print()
    else:
        print(f"Test file not found: {test_file}")
        
        # Test with mock data
        print("Testing with mock external link mapping:")
        mock_mapping = {
            1: "source1.xlsx",
            2: "data_source.xlsx",
            3: "calculation_table.xlsx"
        }
        
        test_formulas = [
            "=[1]Sheet1!A1",
            "=[2]Data!B1",
            "=[3]Table!C1",
            "=SUM([1]Sheet1!A1:[1]Sheet1!A10)",
            "=IF([1]Sheet1!A1>0,[2]Data!B1,0)"
        ]
        
        for formula in test_formulas:
            resolved = resolve_external_references(formula, mock_mapping)
            print(f"  Original: {formula}")
            print(f"  Resolved: {resolved}")
            print()
    
    print("✅ External link mapping test completed")

def test_excel_reading():
    """Test enhanced Excel reading functionality"""
    print("\n" + "=" * 60)
    print("Testing enhanced Excel reading")
    print("=" * 60)
    
    # Test with our test files
    test_files = [
        '/home/runner/work/watchdog/watchdog/test_setup/test_file.xlsx',
        '/home/runner/work/watchdog/watchdog/test_setup/test_external_links.xlsx'
    ]
    
    for test_file in test_files:
        if os.path.exists(test_file):
            print(f"\nTesting file: {os.path.basename(test_file)}")
            try:
                # We need to set up the required config variables
                import sys
                sys.path.append('/home/runner/work/watchdog/watchdog')
                
                # Mock the config variables needed
                import watch
                watch.ENABLE_FAST_MODE = True
                watch.USE_LOCAL_CACHE = False
                watch.CACHE_FOLDER = "/tmp"
                
                # Create a simple copy_to_cache function
                def mock_copy_to_cache(path):
                    return path
                
                # Replace the function temporarily
                original_copy_to_cache = getattr(watch, 'copy_to_cache', None)
                watch.copy_to_cache = mock_copy_to_cache
                
                # Test the enhanced reading
                result = dump_excel_cells_with_timeout(test_file)
                
                print(f"  Worksheets found: {len(result)}")
                for ws_name, ws_data in result.items():
                    print(f"    {ws_name}: {len(ws_data)} cells")
                    for cell_coord, cell_data in list(ws_data.items())[:3]:  # Show first 3 cells
                        print(f"      {cell_coord}: {cell_data}")
                
                # Restore original function
                if original_copy_to_cache:
                    watch.copy_to_cache = original_copy_to_cache
                
            except Exception as e:
                print(f"  Error reading file: {e}")
                import traceback
                traceback.print_exc()
        else:
            print(f"Test file not found: {test_file}")
    
    print("\n✅ Excel reading test completed")

def test_comparison_logic():
    """Test the comparison logic with ArrayFormula objects"""
    print("\n" + "=" * 60)
    print("Testing comparison logic")
    print("=" * 60)
    
    # Simulate the scenario described in the requirements
    print("Simulating ArrayFormula object address change scenario:")
    
    # Create two cells with ArrayFormula objects that have same content but different addresses
    af1 = ArrayFormula(ref="A1:A3")
    af1.text = "SUM(A1:A3)"
    
    af2 = ArrayFormula(ref="A1:A3")
    af2.text = "SUM(A1:A3)"
    
    # Create cell data as it would appear in the watchdog system
    old_cell_data = {
        "formula": "=SUM(A1:A3)",
        "value": serialize_cell_value(af1)
    }
    
    new_cell_data = {
        "formula": "=SUM(A1:A3)",
        "value": serialize_cell_value(af2)
    }
    
    print(f"Old cell data: {old_cell_data}")
    print(f"New cell data: {new_cell_data}")
    print(f"Cell data equal: {old_cell_data == new_cell_data}")
    
    # This should now return True (no change detected) because we're comparing
    # the formula content rather than the object addresses
    
    if old_cell_data == new_cell_data:
        print("✅ SUCCESS: ArrayFormula object address change properly filtered!")
    else:
        print("❌ FAILURE: ArrayFormula object address change not filtered!")
    
    print("\n✅ Comparison logic test completed")

if __name__ == "__main__":
    test_array_formula_filtering()
    test_external_link_mapping()
    test_excel_reading()
    test_comparison_logic()