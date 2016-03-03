Requirements:

Must pass in an excel file

In column A, starting at any row, the following structure must be in place:

MetricID

MetricDataType

1

2

...

n


where the numbers are the test cases.

The script finds the row with "MetricID" and assumes "MetricDataType" is right below it, and the test cases begin below that.


From the cell with "MetricID", the script searches each column until it finds a cell that is type int. It considers this the first metric column.


From the first metric column, the script looks for a blank cell and considers that the last column that needs to be looked at.

	This means that the test grid MUST contain a metric id for each column.

	If there are metrics to be tested that do not yet have a metric id, they must be in columns after all the defined metrics. If there is a column with no metric id, that will be considered the end and no further columns will be considered.
