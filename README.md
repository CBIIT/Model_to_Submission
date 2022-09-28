# Model_to_Submission
This script takes the three files that make a CBIIT data model: model, properties, and terms, and creates a submission workbook with formatting and enumerated drop down menus.

Run the following command in a terminal where R is installed for help.

```
Rscript --vanilla Model_to_Submission.R -h
Usage: Model_to_Submission.R [options]

Model_to_Submission.R version 1.0

Options:
	-m CHARACTER, --model=CHARACTER
		Model file yaml

	-p CHARACTER, --property=CHARACTER
		Model property file yaml

	-t CHARACTER, --terms=CHARACTER
		Model terms file yaml

	-r CHARACTER, --readme=CHARACTER
		README xlsx page (optional)

	-h, --help
		Show this help message and exit
```
