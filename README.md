# Camp Registration Tracker - Simple Working Version

## What This Does

Paste individual rosters → Script detects program/year → Replaces that program/year's data → Tracks cancellations

**Simple. That's it.**

## How It Works

### Structure

**Input:** 51 CSV columns from Daxko roster
**Output:** Master Data with 56 columns (51 CSV + 5 calculated)

**Calculated columns added:**
- `Year` - extracted from Program Start date
- `WeekNumber` - parsed from Instance Name (e.g., "Week 8")
- `SiteDisplay` - cleaned site name (e.g., "NEB", "Animal Camp")
- `DateProcessed` - when first added to Master Data
- `LastUpdated` - when last updated

### Deduplication Key

```
Member ID + Program + Site + Instance Name + Year
```

**Why this works:**
- Same kid in Week 1 and Week 2 = 2 different records ✅
- Same kid in 2025 and 2026 = 2 different records ✅
- Same kid, same week, same year = 1 record (no duplicates) ✅

### Logic Flow

1. **Read paste data** from Paste Report tab (starting at A1)
2. **Detect program and year** from first data row
3. **Add calculated columns** to each row
4. **Read existing Master Data**
5. **Merge intelligently:**
   - Keep all rows from OTHER programs/years
   - For THIS program/year:
     - If record exists in new data → Update it (preserve DateProcessed, update LastUpdated)
     - If record missing from new data → Log as cancellation
   - Add brand new registrations
6. **Write back to Master Data**
7. **Log cancellations** to Cancellations tab (auto-creates if needed)

## Usage

### First Time Setup

1. Create a Google Sheet with a tab named `Paste Report`
2. Go to Extensions → Apps Script
3. Delete any existing code
4. Paste the entire `CampRegistration.gs` file
5. Save (Ctrl+S)
6. Refresh your spreadsheet
7. You'll see a new menu: `📊 Camp Registration`

### Daily/Weekly Use

1. Download roster CSV from Daxko
2. Open CSV, select ALL (Ctrl+A), copy
3. Go to Paste Report tab
4. Click cell A1
5. Paste (Ctrl+V)
6. Click: Camp Registration → Process Updates
7. Confirm the detected program and year
8. Done!

### Multiple Programs

You can paste rosters one at a time:
- Monday: Paste Camp Winnebago 2026 → Process
- Tuesday: Paste Adventure Camps SAY/NEB 2026 → Process
- Wednesday: Paste Good Shepherd 2026 → Process

Master Data will contain ALL programs, properly separated.

### Updating Same Program

If you paste Camp Winnebago 2026 on Monday, then paste it again on Friday:
- Friday's paste REPLACES all Camp Winnebago 2026 data
- Other programs unchanged
- Cancellations logged (kids who were in Monday's data but not Friday's)

## Program Detection

The script auto-detects programs from the "Program" column:

- `2026 Summer Camp - Camp Winnebago` → **Camp Winnebago**
- `2026 Summer Camp - Adventure Camps (SAY/NEB)` → **Adventure Camps (SAY/NEB)**
- `2026 Summer Camp - Adventure Camp (Good Shepherd)` → **Adventure Camp (Good Shepherd)**

Year is extracted from "Program Start" column.

## Site Name Mapping

The script cleans site names for easier analysis:

**Camp Winnebago:**
- `(AC) Animal Camp (Entering 1st - 2nd Grade)` → `Animal Camp`
- `(KC) Kinder Camp (Entering Kindergarten)` → `KinderCamp`
- `(SB) Standing Bear Camp (Entering 3rd - 4th Grade)` → `Standing Bear`
- `(WC) Wilderness Camp (Entering 5th - 6th Grade)` → `Wilderness`
- `(SE) Soaring Eagle Camp (Entering 7th - 8th Grade)` → `Soaring Eagle`

**Adventure Camps:**
- `(ADVCN) Adventure Camp - Northeast Family YMCA` → `NEB`
- `(ADVCS) Adventure Camp - SwedishAmerican YMCA` → `SAY`
- `(GSY) Good Shepherd YMCA` → `Good Shep`

## Cancellation Tracking

When you paste fresh data, the script compares:
- **Old data:** What was in Master Data for this program/year
- **New data:** What you just pasted

**Cancellations = Rows in old data but NOT in new data**

These get logged to the `Cancellations` tab with:
- Date detected
- Member ID, name
- Program, site, instance, year

## What Was Fixed

### Previous Issues:
1. ❌ Data was appending instead of replacing
2. ❌ Columns were duplicating
3. ❌ Cancellation tab never created
4. ❌ Overly complex code

### Now:
1. ✅ Data properly replaces for matching program/year
2. ✅ Clean column structure (51 CSV + 5 calculated, no duplicates)
3. ✅ Cancellation tab auto-creates when cancellations detected
4. ✅ Simple, readable code (~490 lines)

## Key Simplifications

**Before:** Complex record objects, column mapping, dynamic header building
**After:** Simple arrays, fixed column structure, straightforward logic

**Before:** Tried to be too smart about column ordering
**After:** Assumes CSV columns stay consistent (which they do from Daxko)

**Before:** ~800 lines of confusing code
**After:** ~490 lines of clear code

## Testing Checklist

- [ ] Paste first roster → Master Data created
- [ ] Paste same roster again → Data replaced, not appended
- [ ] Remove some rows, paste again → Cancellations logged
- [ ] Paste different program → Both programs in Master Data
- [ ] Paste different year → Both years in Master Data
- [ ] Column count = 56 (51 CSV + 5 calculated)
- [ ] No duplicate columns
- [ ] WeekNumber parsed correctly (e.g., "Week 8")
- [ ] SiteDisplay cleaned correctly (e.g., "NEB", "Animal Camp")

## Troubleshooting

**Error: "Invalid data format"**
- Make sure you paste WITH headers starting at A1
- Headers should include "Member ID" and "Program"

**Data seems wrong**
- Check Extensions → Apps Script → Executions
- View the log for the latest run
- Look for errors

**Cancellations not detected**
- Cancellations only appear when you RE-paste the same program/year
- First paste of a program/year will have 0 cancellations (nothing to compare to)

## Next Steps

Once this is working reliably, you can:
- Build dashboard tabs with formulas
- Connect to Looker Studio
- Add more calculated columns
- Add data validation

But first: **Make sure the basic paste → process → Master Data update works correctly.**

## Column Index Reference

If Daxko changes their CSV format, update these indices at the top of the script:

```javascript
const COL = {
  MEMBER_ID: 1,      // Column B (0-indexed)
  FIRST_NAME: 2,     // Column C
  LAST_NAME: 3,      // Column D
  PROGRAM: 25,       // Column Z
  PROGRAM_START: 27, // Column AB
  SITE: 24,          // Column Y
  INSTANCE: 30       // Column AE
};
```

Count from column A (index 0) to find the right index.
