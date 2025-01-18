# Bank Statement Formatter
Formats bank statements across multiple banks and Splitwise into a format ready for weekly spending analysis in Excel.

#### Steps

> Export date range should buffer 2 days either side of start and end week. Dates in script will filter accurately. Should not export data until 2 days after the final 4+ day week of the month to allow dates to align correctly.

1. Export ANZ Debit
1. Export ANZ Credit
1. Export Splitwise
1. Export Wise
1. Move into correct folders in OneDrive (`C:\Users\Oliver.Sharplin\OneDrive\Documents\Money Stuff\Spending Tracker`)
1. Execute script once for each week covered by the export
    - Example: `python .\spending_formatter.py "WISE" "DEBIT" "CREDIT" "SPLITWISE FOLDER" "YYYY-MM-DD" "YYYY-MM-DD"`
    - `-` as argument will ignore files
1. Copy data from downloads into Spending Tracker tab
1. Copy amounts to NZD converting if necessary (rough currency avg that week)
1. Add tags
1. Create Pivot table
1. Create summary row and enter data (double check for negatives)
1. Review and make comments on areas you need to focus on reducing
1. Once per month, calculate food costs owed to mum and pay her back
    1. Calculate total owed ($5 breakfast, $5 lunch, $10 dinner)
    1. Add food costs that do not fall under mum paying for them to owed amount (such as when on holiday without family)
    1. Calculate total paid amount and make payment - note in description