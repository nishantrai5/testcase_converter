# TestCase Converter Tool

## Introduction

This python script is used to convert testcases written in Excel to TestLink XML format. It is useful when you want to import testcases to TestLink from Excel.
Additionally it provides the option to log templates for the testcases in Excel in Markdown format. Testers can then fill in their execution related information in individual log file and upload them to testlink for a consisten format of logs across multiple testers/teams

# Requirements

- Python 3.6 or above -- [Download](https://www.python.org/downloads/)
- pip   -- [Download](https://pip.pypa.io/en/stable/installing/)
- virtaulenv -- optional -- for better isolation of dependencies -- ```pip install virtualenv```

## Installation

### Clone the repository

```shell
git clone <repository_url>
```

**Folder Struture ::**

```shell
<repository_directory>
├── README.md
├── pip-requirements.txt # contains the list of dependencies to be installed via pip
├── ExcelTemplate.xlsx # Excel template for testcases and column structure
├── xmlTemplate.xml # XML template as used by TestLink for testcases upload
│── MarkdownTemplate.md # Markdown template for logging testcases execution
└── converter.py
```

### Install the dependencies

The Dependices can be installed either directly or in a virtual environment for better isolation of the dependencies

#### For installing directly

```shell
cd <repository_directory>
pip install -r pip-requirements.txt
```

#### For installing in a virtual environment

```shell
cd <repository_directory>
virtualenv venv
source venv/bin/activate
pip install -r pip-requirements.txt
```

## Usage

```shell
Usage: converter.py [OPTIONS]   <file_path> <sheet_name>
      This script converts the testcases in Excel to TestLink XML format.
      Additionally it provides the option to log templates for the testcases in
      Excel in Markdown format. Testers can then fill in their execution related
      information in individual log file and upload them to testlink for a
      consisten format of logs across multiple testers/teams

Options: 
  -m, --template TEXT  Generate Markdown Template to be used for logging testcases execution
                       [markdown]  [default: xml]

If the <sheet_name> is not provided the method does not work even with one sheet in the excel file

Examples : 

python3 converter.py ExcelTemplate.xlsx # generates XML file for uploading to TestLink using the first sheet in the Excel file 

python3 converter.py ExcelTemplate.xlsx Sheet1 # generates XML file for uploading to TestLink with the same name as the input file

python3 converter.py ExcelTemplate.xlsx Sheet1 xmlTemplate.xml # generates XML file for uploading to TestLink with the specified name

python3 converter.py -m ExcelTemplate.xlsx Sheet1 # generates Markdown template for logging 

```

**Examples ::**

```shell
python3 converter.py ExcelTemplate.xlsx # generates XML file for uploading to TestLink using the first sheet in the Excel file 

python3 converter.py ExcelTemplate.xlsx Sheet1 # generates XML file for uploading to TestLink with the same name as the input file

python3 converter.py -m ExcelTemplate.xlsx Sheet1 # generates Markdown template for logging testcases execution
```

## Excel Template Structure 

The Excel template is used to define the structure of the testcases. The script will read the Excel file and generate the XML file for uploading to TestLink. The Excel template should have the following columns:

- **ExternalID** - The External ID of the testcase. This will be used as the external id for the testcase in TestLink for referencing test cases. *This is a optional field.*
- **Name** - The name of the testcase. This will be used as the name of the testcase in TestLink. *This is a mandatory field.*
- **Summary** - The summary of the testcase. This will be used as the summary of the testcase in TestLink. *This is a mandatory field.*
- **PreCondition** - The preconditions of the testcase. This will be used as the preconditions of the testcase in TestLink. *This is a optional field.*
- **Action** - The step/action of the testcase. This will be used as the step - action of the testcase in TestLink. *This is a mandatory field.*
- **ExpectedResults** - The expected result of the testcase step. This will be used as the expected result of the testcase step in TestLink. *This is a mandatory field.*

**Note ::** 

- The Excel template should have the same column names as mentioned above and is case-sensitive. The order of the columns does not matter.

- Additional Columns in the field will be ignored. This allows tester to maintain additional information that may not be need in TestLink.

| ExternalID | Name | Summary     | PreCondition | Action                          | ExpectedResults    |
|------------|------|-------------|--------------|---------------------------------|--------------------|
| I1         | TC1  | Test Case 1 | 1- P1 2- P2  | login                           | user should login  |
|            |      |             |              | add item                        | item added to cart |
|            |      |             |              | logout                          | user should logout |
| I2         | TC2  | Test Case 2 | 1- P1        | Login with Invalid Credtentials | Process should Fail|
|            |      |             |              | Login with Valid Credentials    | Pass               |




