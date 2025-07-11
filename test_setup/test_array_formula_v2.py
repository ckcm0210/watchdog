#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test ArrayFormula handling - understanding constructor
"""

import sys
import os
sys.path.append('/home/runner/work/watchdog/watchdog')

from openpyxl.worksheet.formula import ArrayFormula

def test_array_formula_constructor():
    """Test ArrayFormula constructor and properties"""
    
    print("Testing ArrayFormula constructor...")
    
    # Try different ways to create ArrayFormula
    try:
        # Method 1: Direct construction
        af1 = ArrayFormula()
        print(f"Empty ArrayFormula: {af1}")
        print(f"Has formula attr: {hasattr(af1, 'formula')}")
        print(f"Has text attr: {hasattr(af1, 'text')}")
        
        # Check available attributes
        print(f"ArrayFormula attributes: {dir(af1)}")
        
        # Try to set formula
        if hasattr(af1, 'formula'):
            af1.formula = "SUM(A1:A3)"
            print(f"Set formula: {af1.formula}")
        elif hasattr(af1, 'text'):
            af1.text = "SUM(A1:A3)"
            print(f"Set text: {af1.text}")
        
        # Create another one
        af2 = ArrayFormula()
        if hasattr(af2, 'formula'):
            af2.formula = "SUM(A1:A3)"
        elif hasattr(af2, 'text'):
            af2.text = "SUM(A1:A3)"
        
        print(f"af1 == af2: {af1 == af2}")
        print(f"repr(af1): {repr(af1)}")
        print(f"repr(af2): {repr(af2)}")
        print(f"repr(af1) == repr(af2): {repr(af1) == repr(af2)}")
        
        # Test with different formulas
        af3 = ArrayFormula()
        if hasattr(af3, 'formula'):
            af3.formula = "SUM(B1:B3)"
        elif hasattr(af3, 'text'):
            af3.text = "SUM(B1:B3)"
        
        print(f"af1 formula == af3 formula: {getattr(af1, 'formula', getattr(af1, 'text', '')) == getattr(af3, 'formula', getattr(af3, 'text', ''))}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

def test_improved_serialize():
    """Test improved serialize function"""
    
    print("\n" + "="*60)
    print("Testing improved serialize function")
    print("="*60)
    
    def improved_serialize_cell_value(value):
        """
        Improved serialize function that handles ArrayFormula objects properly
        by comparing their formula content rather than object address
        """
        if value is None:
            return None
        
        # Handle ArrayFormula objects
        if type(value).__name__ == "ArrayFormula":
            # Get the actual formula content, not the object representation
            if hasattr(value, 'formula'):
                return str(value.formula)
            elif hasattr(value, 'text'):
                return str(value.text)
            else:
                # Fallback to string representation but normalize it
                return str(value)
        
        # Handle other objects with formula attribute
        if hasattr(value, 'formula'):
            return str(value.formula)
        
        # Handle standard types
        if isinstance(value, (int, float, str, bool)):
            return value
        
        # Handle datetime (from original code)
        from datetime import datetime
        if isinstance(value, datetime):
            return value.isoformat()
        
        # Default to string representation
        return str(value)
    
    # Test with regular values
    test_values = [
        None,
        "test string",
        123,
        45.67,
        True,
    ]
    
    for val in test_values:
        result = improved_serialize_cell_value(val)
        print(f"Value: {val} -> {result}")
    
    # Test with ArrayFormula if we can create them
    try:
        af1 = ArrayFormula()
        af2 = ArrayFormula()
        
        # Try to simulate different scenarios
        print(f"\nArrayFormula test:")
        print(f"af1: {improved_serialize_cell_value(af1)}")
        print(f"af2: {improved_serialize_cell_value(af2)}")
        print(f"Equal: {improved_serialize_cell_value(af1) == improved_serialize_cell_value(af2)}")
        
    except Exception as e:
        print(f"ArrayFormula test error: {e}")

if __name__ == "__main__":
    test_array_formula_constructor()
    test_improved_serialize()