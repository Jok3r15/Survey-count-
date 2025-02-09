Survey Data Automation
This Python script automates the processing and categorization of customer survey results from an Excel file. It reads the survey data, organizes it by agent name and survey result (Satisfied, Dissatisfied, Neutral), and outputs the results into a new Excel file with categorized sheets, making it easier to analyze survey outcomes.

Note: The data used in this project is false and was generated as a guide. The original data used is sensitive, and for privacy reasons, all names, case numbers, and comments have been altered.

Features
Categorization: Categorizes surveys into Satisfied, Dissatisfied, and Neutral based on survey results.
Survey Data Organization: Stores agent names, case numbers, and comments for each survey.
Alphabetical Sorting: Sorts agent names alphabetically in the output file for better organization.
Excel Output: Creates a new Excel file with the following sheets:
Total Surveys per Agent: Total survey count per agent.
Satisfied Surveys: Surveys marked as "Satisfied" with case numbers and comments.
Dissatisfied Surveys: Surveys marked as "Dissatisfied" with case numbers and comments.
Neutral Surveys: Surveys marked as "Neutral" with case numbers and comments.
Requirements
Python 3.x
openpyxl: Python library for reading and writing Excel files.
Installation

Install the required Python dependencies:

pip install openpyxl
Usage
Place your survey data in an Excel file (e.g., Surveys Week 1 with false data.xlsx) with the following columns:

Agent Name: Name of the survey agent.
Case Number: Case number for the survey.
Comment: The customer's comment regarding the survey.
Survey Result: The result of the survey (i.e., "Satisfied", "Dissatisfied", or "Neutral").
Run the script:

python script_name.py
The script will process the survey data, categorize it, and save the output as a new Excel file with the categorized survey results.

The generated file will be saved with a timestamped name, e.g., Survey_Results_Week_1_Separated_20250209_102030.xlsx.

Example Output
The script generates an Excel file with the following sheets:

Total Surveys per Agent
Satisfied Surveys
Dissatisfied Surveys
Neutral Surveys
Each sheet contains the relevant data, including:

Agent names
Total survey counts
Case numbers and comments (if available) for each survey result category.
License
This project is licensed under the MIT License
