Explanation of Customization Options:

i = 2
This sets the starting row for processing. Change 2 to a different value (e.g., 0 for the first row) if you want to start at a different point.

31 in oSheet.Rows.removeByIndex(i, 31)
Adjust 31 to control how many rows are deleted between kept rows. For example, change it to 15 to keep every 16th row.

Dynamic Last Row Updates
The line lLastRow = oSheet.Rows.getCount() - 1 ensures the macro recalculates the total row count after each deletion. You can comment it out if you're confident your dataset won't change.

Ending Condition in the Loop
The Do While i <= lLastRow loop runs until all deletions are processed. Modify the condition if you'd like to stop processing earlier or based on custom criteria.
