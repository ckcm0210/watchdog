#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test ArrayFormula handling - understanding the real problem
"""

import sys
import os
sys.path.append('/home/runner/work/watchdog/watchdog')

from openpyxl.worksheet.formula import ArrayFormula

def test_array_formula_constructor():
    """Test ArrayFormula constructor and properties"""
    
    print("Testing ArrayFormula constructor...")
    
    try:
        # Create ArrayFormula with ref parameter
        af1 = ArrayFormula(ref="A1:A3")
        print(f"ArrayFormula with ref: {af1}")
        print(f"Ref: {af1.ref}")
        print(f"Has formula attr: {hasattr(af1, 'formula')}")
        print(f"Has text attr: {hasattr(af1, 'text')}")
        
        # Check available attributes
        attrs = [attr for attr in dir(af1) if not attr.startswith('_')]
        print(f"ArrayFormula attributes: {attrs}")
        
        # Create another one with same ref
        af2 = ArrayFormula(ref="A1:A3")
        print(f"af1 == af2: {af1 == af2}")
        print(f"repr(af1): {repr(af1)}")
        print(f"repr(af2): {repr(af2)}")
        print(f"repr(af1) == repr(af2): {repr(af1) == repr(af2)}")
        
        # Test with different ref
        af3 = ArrayFormula(ref="B1:B3")
        print(f"af1.ref == af3.ref: {af1.ref == af3.ref}")
        
        # Try to set formula/text
        if hasattr(af1, 'formula'):
            af1.formula = "SUM(A1:A3)"
            print(f"Set formula on af1: {af1.formula}")
        
        if hasattr(af1, 'text'):
            af1.text = "SUM(A1:A3)"
            print(f"Set text on af1: {af1.text}")
        
        # Test what happens when we get the formula
        for attr in ['formula', 'text']:
            if hasattr(af1, attr):
                val = getattr(af1, attr)
                print(f"af1.{attr}: {val}")
                
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

def demonstrate_array_formula_issue():
    """Demonstrate the actual issue with ArrayFormula object addresses"""
    
    print("\n" + "="*60)
    print("Demonstrating ArrayFormula object address issue")
    print("="*60)
    
    # This simulates what happens when the same Excel file is read multiple times
    # The ArrayFormula objects will have different memory addresses even if they
    # represent the same formula
    
    def simulate_excel_read_1():
        """Simulate reading Excel file - first time"""
        af = ArrayFormula(ref="A1:A3")
        if hasattr(af, 'text'):
            af.text = "SUM(A1:A3)"
        elif hasattr(af, 'formula'):
            af.formula = "SUM(A1:A3)"
        return af
    
    def simulate_excel_read_2():
        """Simulate reading Excel file - second time"""
        af = ArrayFormula(ref="A1:A3")
        if hasattr(af, 'text'):
            af.text = "SUM(A1:A3)"
        elif hasattr(af, 'formula'):
            af.formula = "SUM(A1:A3)"
        return af
    
    # This is what happens in the real scenario
    af1 = simulate_excel_read_1()
    af2 = simulate_excel_read_2()
    
    print(f"First read - ArrayFormula: {repr(af1)}")
    print(f"Second read - ArrayFormula: {repr(af2)}")
    print(f"Objects are equal: {af1 == af2}")
    print(f"Object addresses are same: {id(af1) == id(af2)}")
    print(f"Object reprs are same: {repr(af1) == repr(af2)}")
    
    # Show what the current serialization does
    print("\nCurrent serialization (str()):")
    print(f"str(af1): {str(af1)}")
    print(f"str(af2): {str(af2)}")
    print(f"str(af1) == str(af2): {str(af1) == str(af2)}")
    
    # Show what we want to compare instead
    print("\nWhat we should compare (formula content):")
    formula1 = getattr(af1, 'formula', getattr(af1, 'text', str(af1)))
    formula2 = getattr(af2, 'formula', getattr(af2, 'text', str(af2)))
    print(f"Formula 1: {formula1}")
    print(f"Formula 2: {formula2}")
    print(f"Formula1 == Formula2: {formula1 == formula2}")

if __name__ == "__main__":
    test_array_formula_constructor()
    demonstrate_array_formula_issue()