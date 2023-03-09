# Model_to_Submission
This script takes the three files that make a CBIIT data model: model, properties, and terms, and creates a submission workbook with formatting and enumerated drop down menus.

Run the following command in a terminal where R is installed for help.

```
Rscript --vanilla Model_to_Submission.R -h
Usage: Model_to_Submission.R [options]

Model_to_Submission.R version 2.0.5

This script takes the three files that make a CBIIT data model: model, properties, and terms, and creates a submission workbook with formatting and enumerated drop down menus.

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

An example set of files for this script can be found in the CCDI data model directory: https://github.com/CBIIT/ccdi-model/tree/main/model-desc

```
Rscript --vanilla Model_to_Submission.R -m ccdi-model.yml -p ccdi-model-props.yml -t terms.yaml
```

There is also a set of example files located in the Example_Files directory and the following command can be executed:
```
Rscript --vanilla Model_to_Submission.R -m Example_Files/ccdi-model.yml -p Example_Files/ccdi-model-props.yml -t Example_Files/ccdi-model-props_terms.yaml -r Example_README.xlsx
```
