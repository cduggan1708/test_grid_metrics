Installation:

1. Install git (with Git Bash) - https://git-scm.com/downloads

2. Download the latest Python 3 and pip (same download) (be sure to check off the add Python to your PATH option)

3. Open Git Bash (and use it for all the following commands)
	Type "python --version" to verify the installation was successful

4. Install the openpyxl python library:
	Type "python -m pip install openpyxl", press Enter

5. If you do not have one yet, create a (free) GitHub account - https://github.com

6. Fork my repository so you have your own local copy:
	a. Go to https://github.com/cduggan1708/test_grid_metrics
	b. Click "Fork"
	c. Go back to your GitHub account and verify you have a copy of test_grid_metrics

7. Clone the repo
	a. On your GitHub account, in the forked repo of test_grid_metrics, copy the clone URL (the clipboard icon)
	b. Type "git clone " and paste the cloned URL, press Enter
	c. You now have the forked repo cloned to your host machine




Execution:

From the directory where the git repo is cloned, type the following command (in Git Bash):
	
	python test_grid_metrics.py -h

		This will show you how to use the script

To run it with the test grid template:
	I recommend first opening TestGridTemplate.xlsx and taking note of the sheets (only sheet present is Grid)
	Close that file (this is required before running the script b/c the script needs to be able to write to a new sheet in this file)

	python test_grid_metrics.py -f TestGridTemplate.xlsx

	Once the script has finished executing, open the excel file again and take note of the InsertData sheet

	The first column is memberID, second column is metricID, third column is metric value and fourth column is the query that will be copied into the database to actually insert the data.

	I recommend spot checking the data against the grid to verify it looks accurate before entering it into the DB.




Requirements:

Must pass in an excel file

The test grid sheet must be named 'Grid'

File must not be open while script is executing (cannot write to an open file)

Cells must contain the metric value (that is, not an x or some placeholder)

REFERENCE CODES are not currently handled so DO NOT include them in the excel file (can add them back after the script has finished)

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






USE TestGridTemplate.xlsx AS TEMPLATE
