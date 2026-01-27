# Project Context

## Project Overview
Excel VBA-based orderflow tracking system for SGX stock trading signals. Tracks BULL and BEAR signals across individual stock sheets, compiles them into a Ranking sheet with side-by-side layout, and generates TradingView watchlist strings.

## Current Phase
**COMPLETED**: Ranking Sheet Column Reordering - Moved key columns after Ticker for better visibility

## Completed Work

### Ranking Sheet Restructure (Jan 2026)
**Files Modified:**
- [Module2.bas](c:\projects\orderflow-excel\Module2.bas) - Primary implementation
  - GenerateRankingTable() (lines 1432-1731)
  - WriteQuickRanking() (lines 1234-1537)
  - QuickRankingUpdate() (lines 900-1230)
  - HighlightSignalRow() (lines 1859-1880)
  - SortSignalArray() (lines 1733-1857)
  - BuildTradingViewString() (simplified to read ticker from column 3)
  - LookupTickerCode() (lines 1539-1560)
  - CheckTickerInSheet() (lines 1562-1593)

- [Module3.bas](c:\projects\orderflow-excel\Module3.bas) - Duplicate implementation
  - GenerateRankingTable() (lines 2-286)
  - LookupTickerCode() (duplicate function)
  - CheckTickerInSheet() (duplicate function)

**Changes Implemented:**
1. **Removed columns**: Current_Price, Success_Price, Fail_Price (3 price tracking columns)
2. **Added columns**:
   - Ticker (column 3) - Shows ticker code from Watchlist lookup
   - Bullish (column 8) - Shows "Bullish" if ticker found in Bullish sheet
   - Bearish (column 9) - Shows "Bearish" if ticker found in Bearish sheet
3. **Array structure**: Changed from 8 columns to 9 columns throughout
4. **Column layout**:
   - BULL section: A-I (was A-H)
   - Spacer: J (was I)
   - BEAR section: K-S (was J-Q)

**New Column Structure (Jan 2026 - Reordered):**
```
BULL (A-I):          BEAR (K-S):
1. Rank              11. Rank
2. Stock             12. Stock
3. Ticker            13. Ticker
4. Entry_Price       14. Entry_Price
5. Accel_Count       15. Accel_Count
6. Bullish           16. Bearish      <- Note: BEAR shows Bearish first
7. Bearish           17. Bullish      <- Note: BEAR shows Bullish second
8. Signal_Type       18. Signal_Type
9. Signal_Status     19. Signal_Status
```

**Column Reordering Rationale:**
- Entry_Price and Accel_Count moved right after Ticker for quick reference
- Bullish/Bearish flags placed prominently before Signal_Type/Status
- BEAR section shows Bearish before Bullish (most relevant flag first for each section)

**Helper Functions:**
- `LookupTickerCode(stockName)`: Maps stock name to ticker code via Watchlist sheet (Column C → Column D)
- `CheckTickerInSheet(ticker, sheetName)`: Checks if ticker exists in comma-separated string in target sheet cell A1

**Data Sources:**
- **Watchlist sheet**: Column C = Stock names, Column D = Ticker codes
- **Bullish sheet**: Cell A1 contains comma-separated tickers (e.g., "5E2, 5TP, AU8U, S56...")
- **Bearish sheet**: Cell A1 contains comma-separated tickers (same format)

## In Progress
None - performance optimization completed successfully

### Performance Optimization (Jan 2026)
**Files Modified:**
- [Module2.bas](Module2.bas) - Major performance rewrite

**Changes Implemented:**
1. **Binary Search for Velocity** (lines 181-270)
   - Added `BinarySearchTime()` helper function
   - Added `CalcVelocityFromArrays()` optimized function
   - Reduces O(n²) backward walk to O(n log n) binary search

2. **Bulk Array Processing in ProcessSingleStock** (lines 543-850)
   - Complete rewrite to read all data into memory arrays in ONE read
   - Process all calculations (E-S columns) in memory
   - Write results back in 4 bulk writes instead of ~60K individual writes
   - ~10-20x speedup for large datasets

3. **Optimized Median Calculation** (lines 560-610)
   - Added `CalculateMedianVolumeOptimized()` using Excel's built-in Median
   - O(n) instead of O(n²) bubble sort

4. **Dictionary Caching for Lookups** (lines 175-290)
   - Added `InitializeLookupCaches()` - loads Watchlist/Bullish/Bearish into Dictionaries
   - Added `LookupTickerCodeCached()` - O(1) ticker lookup
   - Added `IsBullishCached()` / `IsBearishCached()` - O(1) flag checks
   - Added `ClearLookupCaches()` - cleanup

5. **Bulk Writes for Ranking Output**
   - Updated `GenerateRankingTable()` to use bulk array writes
   - Updated `WriteQuickRanking()` to use bulk array writes

**Expected Performance:**
- 10K rows processing: 60-120 sec → 5-10 sec (~10-20x faster)
- End-of-day ranking: 30+ sec → 2-5 sec

## Next Steps
1. Monitor for any edge cases during production use
2. Consider adding error handling if Bullish/Bearish sheets don't exist
3. Update Watchlist sheet with missing ticker entries (some stocks showing full names in TradingView string)

## Blockers
None

## Key Architecture Decisions

### Ticker Column Addition (Jan 2026)
**Decision**: Add visible Ticker column instead of doing repeated Watchlist lookups
**Rationale**:
- Clearer for users (shows both stock name and ticker code)
- Better performance (lookup once during data collection, not on every access)
- Easier to debug (visible ticker codes help verify Bullish/Bearish matching)
- Simplifies BuildTradingViewString() - reads directly from column 3

### 9-Column Array Structure
**Decision**: Use fixed 9-column arrays for BULL and BEAR data
**Rationale**:
- Consistent structure across all functions
- Simplifies sorting and data manipulation
- Clear column mapping reduces index errors
- Room for future expansion if needed

### Side-by-Side Layout
**Decision**: BULL (A-I) | Spacer (J) | BEAR (K-S)
**Rationale**:
- Easy visual comparison of BULL vs BEAR signals
- Batch timestamps aligned for temporal context
- TradingView strings for quick watchlist import
- Separate sections allow independent sorting

### Lookup Strategy
**Decision**: Stock Name → Ticker Code → Bullish/Bearish check
**Rationale**:
- Sheet names are stock names (e.g., "CENT ACCOM REIT")
- Bullish/Bearish lists use ticker codes (e.g., "C38U", "5E2")
- Two-step lookup required: stockName → tickerCode → flag
- Fallback: If Watchlist lookup fails, use stock name as ticker

## Technical Debt
- Module3.bas appears to be a duplicate/older version of Module2 functions - confirm if still actively used
- Some stocks missing from Watchlist sheet (causes stock names to appear in TradingView output instead of tickers)
- No error handling if Bullish/Bearish sheets don't exist (low priority - user creates them manually)

## Notes
- All changes tested and verified through multiple iterations
- Fixed subscript out of range errors by ensuring consistent 9-column structure across all functions
- Color scheme maintained: Light green (Bullish), Light red (Bearish), Light blue (Active), matching existing design
- SortSignalArray updated to use Signal_Status at column 5 (was column 4) and Accel_Count at column 6 (was column 5)

### Column Reordering Update (Jan 2026)
- Reordered columns to show Entry_Price, Accel_Count, Bullish/Bearish right after Ticker
- Updated `HighlightSignalRow()` in both Module2.bas and Module3.bas with new column positions:
  - BULL: Bullish=col 6, Bearish=col 7
  - BEAR: Bearish=col 16, Bullish=col 17
- Data array indices unchanged (Bullish=8, Bearish=9) - only Excel column mapping changed
- Highlighting logic preserved: checks cell VALUE ("Bullish"/"Bearish") at new column positions
