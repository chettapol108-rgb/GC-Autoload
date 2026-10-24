Option Explicit

Public Const GC_AUTOLOAD_VERSION As String = "3.3"
Public Const GC_AUTOLOAD_RELEASE_DATE As String = "2026-10-07"
'==============================================================================
' GC AutoLoad
'==============================================================================
'
' VERSION HISTORY
' v3.3 - 2026-10-07
' - Added SKU quantity validation between Raw Input and the Qty column.
' - Added detection for quantity mismatches in explicit "*Npcs" SKU declarations.
' - Improved input validation to prevent incorrect quantities from being translated.
' - Added clearer error messages when Raw Input quantity does not match order quantity.
' - Improved validation for duplicate Order ID + SKU rows by comparing against total order quantity.
' v3.2 - 2026-09-04
' - Added full Unicode support for XML export.
' - Fixed Korean, Chinese, Japanese, and other Unicode text being removed during XML export.
' - Improved XML text cleaning to preserve valid Unicode characters.
' - Added support for emoji and extended Unicode characters in XML output.
' - Improved XML special-character handling for &, <, >, quotes, and apostrophes.
' - Improved input validation using SKU, Order ID, and quantity.
' - Added support for grouping duplicate Order ID + SKU rows and validating total quantity.
' - Improved translation error messages and highlighted invalid input rows.
' - Standardized quantity input using "p" and "pcs".
'
' v3.1 - 2026-08-24
' - Added GC AutoLoad version information to Ribbon.
' - Added support for "pcs" quantity input.
' - Normalized "pcs" to internal "p" format.
'
' v3.0 - 2026-08-24
' - Converted GC AutoLoad into a centralized Excel Add-in (.xlam).
' - Added custom Ribbon.
' - Centralized shared AutoLoad functions.
'
' v2.4.2
' - Previous standalone workbook version.
'
'==============================================================================
