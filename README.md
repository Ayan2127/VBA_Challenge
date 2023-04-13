# VBA_Challenge Logistics

Includes columns: Ticker Symbol, Yearly Change, Percent Change, Total Stock Volume

Values for above ^^ columns generated using for loops

If...Then...End if statements for conditional formatting

For the Excel sheet to loop successfully, logic must flow (i.e, conditional formatting loop being below the summary table loop)

Yearly change and percent change conditional formatting correspond with one another (i.e., if the yearly change value is negative, then the percent change value is also negative, resulting in the same color code)

The for each ws in worksheets function enables VBA to run all three sheets in one go 

Functions WorksheetFunction.Max & WorksheetFunction.Min calculate greatest % increase, decrease, and greatest volume

***Min/Max VBA functions source: https://www.wallstreetmojo.com/vba-max/

Corresponding ticker for greatest % increase, decrease, and greatest volume populated using If...Then...Elseif statements

Helpful Excel activities that guided VBA logic for this week's challenge: credit_charges, lotto_numbers, & census_data
