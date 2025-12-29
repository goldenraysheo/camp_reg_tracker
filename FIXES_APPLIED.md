# Apps Script Fixes - Camp Registration Tracker

## Issues Fixed

### 1. Data Appending Instead of Replacing ✅

**Problem:**
- Script was duplicating data instead of replacing existing records
- Column index mapping was unreliable when paste data had different column order
- Composite keys failed to match between new and existing data

**Solution:**
- Implemented normalized record objects instead of relying on array indices
- Created proper column mapping system that handles varying column orders
- Records are now compared using structured objects, not raw arrays
- Clear separation between "update existing" and "add new" logic

### 2. Cancellation Tab Not Creating ✅

**Problem:**
- Cancellation detection depended on broken deduplication logic
- When keys didn't match, no cancellations were detected
- Sheet creation logic was correct, but never triggered

**Solution:**
- Fixed deduplication ensures cancellations are properly detected
- Cancellations are now logged whenever a record exists in Master Data but not in new paste data
- Sheet is created automatically on first cancellation detection

### 3. Overcomplexity ✅

**Problem:**
- Mixed concerns between data structure and processing logic
- Column indices used inconsistently
- Hard to understand data flow

**Solution:**
- Clear separation of concerns:
  - `buildRecord()` - normalize incoming data
  - `parseRecord()` - parse existing data
  - `mergeRecords()` - intelligent merge logic
  - `writeMasterData()` - output formatting
- Data flows through normalized record objects
- Easier to debug and maintain

## Key Improvements

### Data Handling
```javascript
// OLD: Direct array indexing (fragile)
const key = `${row[memberIdCol]}|${row[programCol]}|...`;

// NEW: Normalized record objects (robust)
const record = {
  memberId: '...',
  program: '...',
  originalData: {...}
};
const key = buildKey(record);
```

### Merge Logic
```javascript
// NEW: Clear 3-step process
1. Keep records from other programs/years unchanged
2. Update or detect cancellations for matching program/year
3. Add brand new records
```

### Column Mapping
```javascript
// NEW: Dynamic column mapping handles any order
const columnMap = buildColumnMap(headers);
const value = row[columnMap['Member ID']]; // Always correct
```

## How It Works Now

### Processing Flow

1. **Extract & Validate**
   - Read paste data
   - Validate required columns exist
   - Extract program/year metadata

2. **Normalize New Data**
   - Convert each row to record object
   - Calculate derived fields (WeekNumber, SiteDisplay)
   - Store all original data

3. **Load Existing Data**
   - Read Master Data sheet
   - Parse into record objects
   - Build comparison map

4. **Intelligent Merge**
   - **Other Programs/Years:** Keep unchanged
   - **Matching Program/Year:**
     - If record exists in new data → Update it
     - If record missing from new data → Mark as cancellation
   - **Brand New Records:** Add them

5. **Write Results**
   - Clear Master Data
   - Write merged records
   - Log cancellations to separate sheet

### Deduplication Key

Records are uniquely identified by:
```
Member ID | Program | Site | Instance Name | Year
```

This ensures that:
- Same person in different weeks = different records
- Same person in different years = different records
- Same person in same week/year = ONE record (updated, not duplicated)

## Testing Checklist

Test these scenarios to verify fixes:

- [ ] **Initial Load:** Paste data into empty Master Data
- [ ] **Re-paste Same Data:** Should update LastUpdated, not duplicate
- [ ] **Paste with Cancellations:** Some registrations removed from new data
  - Check Cancellations tab is created
  - Verify cancelled records are logged
- [ ] **Paste Different Program:** Other program data unchanged
- [ ] **Paste Different Year:** Previous year data preserved
- [ ] **Column Order Change:** Paste data with columns in different order
- [ ] **Partial Update:** Paste data for one program, others unchanged

## Migration Instructions

### Option 1: Clean Start (Recommended)
1. Backup your current spreadsheet
2. Open Apps Script editor (Extensions > Apps Script)
3. Delete all existing code
4. Paste new code from `CampRegistration.gs`
5. Save (Ctrl+S)
6. Refresh spreadsheet
7. Test with sample data

### Option 2: Keep Existing Data
1. Backup your spreadsheet
2. Replace script code as above
3. Existing Master Data will be processed correctly
4. First run will normalize data structure

## Support

### Common Issues

**"Missing required columns" error**
- Ensure you copied ALL data including headers
- Check column names match exactly

**Cancellations not detected**
- Ensure you're pasting FULL roster, not incremental updates
- System compares current paste vs. existing Master Data

**Duplicates appearing**
- Should not happen with new code
- If it does, check that Member ID values are consistent (no extra spaces)

### Debug Mode

To see detailed logs:
1. Apps Script editor > Executions
2. View execution log for latest run
3. Check for errors or warnings

## Technical Notes

### Why Record Objects?
- Immune to column reordering
- Easier to debug (named fields vs. indices)
- Clear separation of original vs. calculated data
- Type-safe comparisons

### Performance
- Optimized for typical camp sizes (1000-5000 registrations)
- Batch writes to minimize API calls
- Single clear + single write per update

### Data Integrity
- Preserves `DateProcessed` on updates (shows original registration date)
- Updates `LastUpdated` on every change
- Never modifies data from other programs/years
- All original CSV columns preserved
