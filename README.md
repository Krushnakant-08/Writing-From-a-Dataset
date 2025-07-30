# Writing From a Dataset

This project provides a Python script to automatically generate personalized Microsoft Word documents from data stored in an Excel spreadsheet. The script prompts the user for a student's name, finds the corresponding data in the Excel file, and populates a Word template to create a unique document for that student.

## Prerequisites

Before running the script, you need to have Python 3 installed, along with the following libraries:

*   pandas
*   openpyxl
*   docxtpl

## Installation

1.  Clone the repository to your local machine:
    ```sh
    git clone https://github.com/Krushnakant-08/Writing-From-a-Dataset.git
    ```
2.  Navigate to the project directory:
    ```sh
    cd Writing-From-a-Dataset
    ```
3.  Install the required packages using pip:
    ```sh
    pip install pandas openpyxl docxtpl
    ```

## Usage

1.  **Update the Data Source**: Populate the `Sample.xlsx` file with your student data. Ensure the column headers match those used in the script: `Name`, `MotherName`, `StudentID`, `Gen_Reg_no`, `Class`, `DOB`.

2.  **Customize the Template**: The `student.docx` file is a Word template. You can edit this file to change the layout and formatting. The script uses Jinja2-style placeholders (e.g., `{{Name}}`, `{{StudentID}}`) which correspond to the column headers in `Sample.xlsx`.

3.  **Run the Script**: Execute the script from your terminal.
    ```sh
    python script.py
    ```

4.  **Enter Student Name**: The script will prompt you to enter the name of the student for whom you want to generate a document.
    ```
    Enter Name of Student :
    ```

5.  **Check the Output**: If a student with the entered name is found, a new Word document named `Reg_{StudentID}.docx` will be saved in the root directory of the project. For example, `Reg_101.docx`.

## File Descriptions

*   `script.py`: The core Python script that reads the data, processes the template, and generates the output document.
*   `Sample.xlsx`: The Excel spreadsheet that serves as the database for student information.
*   `student.docx`: The Microsoft Word template file containing placeholders for the data.
