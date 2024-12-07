# bank-statement-formatter
Formats bank statements across multiple banks and Splitwise into a format ready for weekly spending analysis in Excel.

Steps

1. Export ANZ Debit
1. Export ANZ Credit
1. Export SplitWise
1. Export Wise
1. Move into correct folders in OneDrive
1. Execute script with parameters (add extra 2 days onto first start date from export due to transaction time changes)
1. Copy data from downloads into Spending Tracker tab
1. Check for date overlap due to added 2 days and delete rows if necessary
1. Copy amounts to NZD converting if necessary (rough currency avg that week)
1. Add tags
1. Create Pivot table
1. Create summary row and enter data
1. Review and make comments on areas you need to focus on reducing
1. Once per month, calculate food costs owed to mum and pay her back
