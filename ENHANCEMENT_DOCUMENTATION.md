# Enhanced Watchdog Implementation

## Overview

This implementation adds two key enhancements to the watchdog Excel monitoring system:

1. **ArrayFormula Change Filtering** - Filters out false positives when ArrayFormula objects have different memory addresses but identical content
2. **External Link Mapping** - Resolves `[n]Table!` references to show actual file paths in reports

## Requirements Implementation

### 1. ArrayFormula Change Filtering (需求 1)

**Problem**: When watchdog compares Excel cells, ArrayFormula objects with identical content but different memory addresses were being detected as changes.

**Solution**: Enhanced `serialize_cell_value()` function to extract actual formula content from ArrayFormula objects instead of using their string representation (which includes memory addresses).

**Key Changes**:
- Modified `serialize_cell_value()` to handle ArrayFormula objects properly
- Added logic to extract `text` or `formula` attributes from ArrayFormula objects
- Updated `print_cell_changes_summary()` to show filtering information

**Code Changes**:
```python
def serialize_cell_value(value):
    # ... existing code ...
    # 處理 ArrayFormula 對象 - 比較公式內容而非物件地址
    elif type(value).__name__ == "ArrayFormula":
        # 取得實際公式內容，避免物件地址差異導致的誤判
        if hasattr(value, 'text'):
            return str(value.text)
        elif hasattr(value, 'formula'):
            return str(value.formula)
        else:
            return str(value)
```

### 2. External Link Mapping (需求 2)

**Problem**: Excel formulas with `[n]Table!` external references were not showing the actual source file paths.

**Solution**: Added ZIP/XML parsing to extract external link mappings from Excel files and resolve `[n]` references to actual file paths.

**Key Changes**:
- Added `extract_external_links()` function to parse Excel ZIP structure
- Added `resolve_external_references()` function to resolve `[n]Table!` references
- Enhanced `dump_excel_cells_with_timeout()` to apply external link resolution
- Uses XML parsing of `/xl/externalLinks/externalLink*.xml` and `/xl/_rels/workbook.xml.rels`

**Code Changes**:
```python
def extract_external_links(excel_file_path):
    """從 Excel 檔案中提取外部連結映射"""
    # ZIP parsing logic to extract external link mappings
    
def resolve_external_references(formula, external_link_mapping):
    """使用外部連結映射解析公式中的 [n]Table! 參照"""
    # Regex-based formula resolution
```

## Testing

Comprehensive test suite includes:

1. **ArrayFormula Tests**: Verify that objects with same content but different addresses are treated as equal
2. **External Link Tests**: Test formula resolution with mock external link mappings
3. **Integration Tests**: Full watchdog scenario simulation
4. **Comparison Tests**: Verify change detection works correctly

## Files Modified

### Primary Changes:
- `watch.py`: Enhanced with ArrayFormula filtering and external link mapping

### Test Files Added:
- `test_setup/test_enhanced_functionality.py`: Core functionality tests
- `test_setup/test_integration.py`: Integration tests
- `test_setup/test_array_formula_v3.py`: ArrayFormula-specific tests
- `test_setup/create_test_excel.py`: Test file creation utilities

## Usage

The enhanced watchdog works identically to the original, but now:

1. **ArrayFormula Changes**: Will not trigger false positives when only object addresses change
2. **External Links**: Will show resolved file paths in reports like:
   ```
   [公式] '=[1]Sheet1!A1' -> '=[source_file.xlsx]Sheet1!A1'
   ```

## Backward Compatibility

All changes are backward compatible:
- Existing functionality remains unchanged
- New features only activate when relevant data is present
- No breaking changes to existing APIs or configuration

## Performance Impact

Minimal performance impact:
- ArrayFormula filtering: Near-zero overhead (only affects comparison logic)
- External link mapping: Small overhead during Excel reading (only when external links exist)
- All enhancements are optional and only activate when needed

## Example Output

### ArrayFormula Filtering:
```
  變更 cell 數量：5
  已過濾 ArrayFormula 物件地址變更：2 個
    [Sheet1] A1:
        [公式] '=SUM(A1:A3)' -> '=SUM(A1:A3)'
        [值] 'SUM(A1:A3)' -> 'SUM(A1:A3)'
```

### External Link Resolution:
```
   🔗 發現外部連結映射: {1: 'source1.xlsx', 2: 'data_table.xlsx'}
    [Sheet1] B1:
        [公式] '=[1]Sheet1!A1' -> '=[source1.xlsx]Sheet1!A1'
        [值] 'external_value' -> 'external_value'
```

## Notes

1. External link mapping requires actual Excel files with real external links to be fully effective
2. The implementation handles edge cases gracefully with proper error handling
3. All debug output includes clear indicators when enhancements are active
4. Test files demonstrate both features working correctly