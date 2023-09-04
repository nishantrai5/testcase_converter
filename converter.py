# Converter Class for conveting test cases saved in a an excel format to xml and markdown format
# The class has three methods:    1. convert_to_xml   2. convert_to_excel   3. convert_to_markdown

####### Excel File Format represented as a markdown table ########

# | UseCase | ExternalID | Name | Summary     | PreCondition | Action                          | ExpectedResults    |
# |---------|------------|------|-------------|--------------|---------------------------------|--------------------|
# | U1      | I1         | TC1  | Test Case 1 | 1- P1 2- P2  | login                           | user should login  |
# |         |            |      |             |              | add item                        | item added to cart |
# |         |            |      |             |              | logout                          | user should logout |
# | U2      | I2         | TC2  | Test Case 2 | 1- P1        | Login with Invalid Credtentials | Correct            |
# |         |            |      |             |              | Login with Valid Credentials    | Invalid            |

####### XML File Format ########

# <?xml version="1.0" encoding="UTF-8"?>
# <testcase name="tc1">
# 	<version><![CDATA[1]]></version>
# 	<summary><![CDATA[<p>demo tc</p>
# ]]></summary>
# 	<preconditions><![CDATA[<p>dinaind</p>
# ]]></preconditions>
# 	<execution_type><![CDATA[1]]></execution_type>
# </testcase>
# </testcases>
# <testcase name="tc2">
# 	<version><![CDATA[1]]></version>
# 	<summary><![CDATA[<p>demo tc</p>
# ]]></summary>
# 	<preconditions><![CDATA[<p>dinaind</p>
# ]]></preconditions>
# 	<execution_type><![CDATA[1]]></execution_type>
# </testcase>
# <testcases>
# <testcase name="tc3">
# 	<version><![CDATA[1]]></version>
# 	<summary><![CDATA[<p>demo tc</p>
# ]]></summary>
# 	<preconditions><![CDATA[<p>dinaind</p>
# ]]></preconditions>
# 	<execution_type><![CDATA[1]]></execution_type>
# <steps>
# <step>
# 	<step_number><![CDATA[1]]></step_number>
# 	<actions><![CDATA[<p>demo step</p>
# ]]></actions>
# 	<expectedresults><![CDATA[<p>demo step</p>
# ]]></expectedresults>
# 	<execution_type><![CDATA[1]]></execution_type>
# </step>
# </steps>
# </testcase>

####### Markdown File Format ########

# # UseCase: U1
# ## TC1
# ### Summary
# Test Case 1
# ### PreCondition
# 1- P1
# 2- P2
# ### Step - 1 | Action: login | Expected Result: user should login
# ### Step - 2 | Action: add item | Expected Result: item added to cart
# ### Step - 3 | Action: logout | Expected Result: user should logout
# ## TC2
# ### Summary
# Test Case 2
# ### PreCondition
# 1- P1
# ### Step - 1 | Action: Login with Invalid Credtentials | Expected Result: Correct
# ### Step - 2 | Action: Login with Valid Credentials | Expected Result: Invalid


import pandas as pd
from xml.etree.ElementTree import Element, SubElement, tostring, ElementTree
from xml.dom import minidom
import sys


class Converter:
    def __init__(self) -> None:
        self.auto_number_steps = False

    def convert_to_xml(self, excel_file_path, sheet_name):
        # Load the excel file
        try:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)

        # Create the root element
        root = Element("testcases")
        current_testcase = None
        test_action_count = 0

        # Check is the External ID column is present
        if "ExternalID" in df.columns:
            add_enternal_id = True

        # Iterate over the rows of the excel file
        for _, row in df.iterrows():
            if pd.notnull(row["Name"]):
                # Create a new testcase element
                current_testcase = SubElement(root, "testcase", name=row["Name"])
                current_testcase.set("name", row["Name"])
                test_action_count = 0
                # Create the version element
                version = SubElement(current_testcase, "version")
                version.text = "1"
                # Create the summary element
                summary = SubElement(current_testcase, "summary")
                summary.text = self._format_text(row["Summary"])
                # Create the preconditions element
                preconditions = SubElement(current_testcase, "preconditions")
                preconditions.text = self._format_text(row["PreCondition"])
                # Create the execution_type element
                execution_type = SubElement(current_testcase, "execution_type")
                execution_type.text = "1"
                # Create the externalid element
                if add_enternal_id:
                    externalid = SubElement(current_testcase, "externalid")
                    externalid.text = row["ExternalID"]
                # Create the steps element
                steps = SubElement(current_testcase, "steps")
            if (
                pd.notnull(row["Action"])
                and pd.notnull(row["ExpectedResults"])
                and current_testcase is not None
            ):
                # Create the step element
                step = SubElement(steps, "step")
                # Create the step_number element
                step_number = SubElement(step, "step_number")
                step_number.text = str(test_action_count + 1)
                # Create the actions element
                actions = SubElement(step, "actions")
                actions.text = self._format_text(row["Action"])
                # Create the expectedresults element
                expectedresults = SubElement(step, "expectedresults")
                expectedresults.text = self._format_text(row["ExpectedResults"])
                # Create the execution_type element
                execution_type = SubElement(step, "execution_type")
                execution_type.text = "1"
                test_action_count += 1

        # Convert the xml to a pretty string
        xmlstr = minidom.parseString(tostring(root)).toprettyxml(indent="   ")

        # Write the xml to a file
        excel_file_name = excel_file_path.split("/")[-1].split(".")[0]
        serial_number = 1
        while True:
            try:
                with open(
                    f"{excel_file_name}_{sheet_name}_{serial_number}.xml", "r"
                ) as f:
                    serial_number += 1
            except FileNotFoundError:
                break
        with open(f"{excel_file_name}_{sheet_name}_{serial_number}.xml", "w") as f:
            f.write(xmlstr)
        # Replace all the &lt; and &gt; with < and > in the output file
        with open(f"{excel_file_name}_{sheet_name}_{serial_number}.xml", "r") as f:
            xml_string = f.read()
        xml_string = xml_string.replace("&lt;", "<")
        xml_string = xml_string.replace("&gt;", ">")
        with open(f"{excel_file_name}_{sheet_name}_{serial_number}.xml", "w") as f:
            f.write(xml_string)
        print(f"XML file saved as {excel_file_name}_{sheet_name}_{serial_number}.xml")

    def convert_to_excel(self, xml_file_path):
        # Load the xml file
        try:
            tree = ElementTree()
            tree.parse(xml_file_path)
        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)

        # Create the dataframe
        df = pd.DataFrame(
            columns=[
                "UseCase",
                "ExternalID",
                "Name",
                "Summary",
                "PreCondition",
                "Action",
                "ExpectedResults",
            ]
        )

        # Iterate over the testcases
        for testcase in tree.iter("testcase"):
            # Create a new row
            row = {}
            row["UseCase"] = testcase.get("name")
            row["ExternalID"] = testcase.get("name")
            row["Name"] = testcase.get("name")
            row["Summary"] = testcase.find("summary").text
            row["PreCondition"] = testcase.find("preconditions").text
            # If not Steps in the XML then move to the next testcase
            if testcase.find("steps") is None:
                df = df.append(row, ignore_index=True)
                continue
            row["Action"] = ""
            row["ExpectedResults"] = ""
            for step in testcase.iter("step"):
                row["Action"] = step.find("actions").text
                row["ExpectedResults"] = step.find("expectedresults").text
                # Append the row to the dataframe
                df = df.append(row, ignore_index=True)

        # Write the dataframe to an excel file
        xml_file_name = xml_file_path.split("/")[-1].split(".")[0]
        serial_number = 1
        while True:
            try:
                with open(f"{xml_file_name}_{serial_number}.xlsx", "r") as f:
                    serial_number += 1
            except FileNotFoundError:
                break
        df.to_excel(f"{xml_file_name}_{serial_number}.xlsx", index=False)

    def convert_to_markdown(self, excel_file_path, sheet_name):
        # Load the excel file
        try:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)

        # Create the markdown string
        markdown_string = ""
        current_usecase = None

        # Iterate over the rows of the excel file
        for index, row in df.iterrows():
            if pd.notnull(row["UseCase"]):
                if current_usecase is None or current_usecase != row["UseCase"]:
                    markdown_string += f"\n# UseCase: {row['UseCase']}\n"
                    current_usecase = row["UseCase"]
            if pd.notnull(row["Name"]):
                markdown_string += f"\n## {row['Name']}\n"
                step_counter = 1
            if pd.notnull(row["Summary"]):
                markdown_string += f"\nSummary : \n\n{row['Summary']}\n"
            if pd.notnull(row["PreCondition"]):
                markdown_string += f"\nPre Condition : \n\n{row['PreCondition']}\n"
            if pd.notnull(row["Action"]) and pd.notnull(row["ExpectedResults"]):
                markdown_string += f"\n### Step - {step_counter} | Action: {row['Action']} | Expected Result: {row['ExpectedResults']}\n"
                step_counter += 1
            markdown_string += "\n"

        # Write the markdown string to a file
        excel_file_name = excel_file_path.split("/")[-1].split(".")[0]
        serial_number = 1
        while True:
            try:
                with open(
                    f"{excel_file_name}_{sheet_name}_{serial_number}.md", "r"
                ) as f:
                    serial_number += 1
            except FileNotFoundError:
                break
        with open(f"{excel_file_name}_{sheet_name}_{serial_number}.md", "w") as f:
            f.write(markdown_string)

    def _format_text(self, text):
        if pd.notna(text):
            return "<![CDATA[<p>" + text.replace("\n", "<br>") + "</p>]]>"
        return ""


if __name__ == "__main__":
    converter = Converter()
    if len(sys.argv) == 4:
        # If user set the flag -m then generate the markdown file
        if sys.argv[1] == "-m":
            excel_file_path = sys.argv[2]
            sheet_name = sys.argv[3]
            converter.convert_to_markdown(excel_file_path, sheet_name)
        # If user set the flag -x then generate the xml file
        elif sys.argv[1] == "-x":
            excel_file_path = sys.argv[2]
            sheet_name = sys.argv[3]
            converter.convert_to_xml(excel_file_path, sheet_name)
    elif len(sys.argv) == 2:
        xml_file_path = sys.argv[1]
        converter.convert_to_excel(xml_file_path)
    else:
        print(
            "Invalid Input Format : python converter.py -m|x <excel_file_path> <sheet_name> or python converter.py <xml_file_path>"
        )
