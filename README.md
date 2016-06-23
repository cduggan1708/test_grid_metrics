Installation:

1. Install git (with Git Bash) - https://git-scm.com/downloads

2. Download the latest Python 3 and pip (same download) (be sure to check off the add Python to your PATH option)

3. Open Git Bash (and use it for all the following commands)
	Type "python --version" to verify the installation was successful

4. Install the openpyxl python library:
	Type "python -m pip install openpyxl", press Enter

5. Clone the repo from https://provanthealth.visualstudio.com/Provant/_git/QA 




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
		You will need the variable @today and @yesterday at the top of the query:
			DECLARE @today VARCHAR(100)
			SET @today = convert(date, CURRENT_TIMESTAMP)

			DECLARE @yesterday VARCHAR(100)
			SET @yesterday = convert(date,DATEADD(d,-1,GETDATE()))


To run with your own excel file, either copy the excel file into this directory or give the full path to the -f argument


Received Date:

If Received Date is applicable to your test cases, pass in -r flag to get the RDC-specific stored procedure execute statements.
	For now the following data is hardcoded and will need to be manually replaced:
		@clientId=10030			# to get the correct id use getProgramAndClientIDs.sql
		@formTemplateId=175		# to get the correct formTemplateId use getClientHSRFFormTemplate.sql



Requirements:

Must pass in an excel file

The test grid sheet must be named 'Grid'

File must not be open while script is executing (cannot write to an open file)

Cells must contain the metric value (that is, not an x or some placeholder)

Reference Codes can be entered in the test grid as the string code (i.e. DNP, N/A) and the script will translate that to the correct reference code id to insert.
Any metric value that is a string will be considered a reference code because other than reference codes, all metric values should be numbers (i.e. float could be 199, boolean is 0 or 1, enum could be 229).
Due to the fact that a common test case will be to have a reference code and an actual metric (number) for a given member, the script sets the reference code query to have @yesterday for measured, created.
	This is necessary to create separate records in metricvalues table (if same measured/create, stored procedure will put reference code in same metricvalue record as float/bool/enum, which we do not want for testing).

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
