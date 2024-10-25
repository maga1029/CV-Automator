# Instructions

1)  Select the .xlsx file that will be read. It has to have the correct format.
2)  Select the destination folder.
3)  Select the "No spaces" button if you want no spaces between paragraphs.
4)  Select the "To PDF" button if you want to convert the .docx format to a .pdf file too.
5)  Wait for the application to work. The window will freeze and won't accept any output. After a while, a message window will appear showing that the process has ended. You may continue to convert CV's or close the application. 

## Excel format

To correctly use the given format for CV creation in an Excel sheet, follow these instructions:

### Sheet Layout:
- Column A, Section: Describes the specific section of the CV. It won't affect the reading of the file if the cells are erased or changed.
- Column B, Code: Defines the hierarchical position and subcategories within a section.
- Column C, Included: Input 1 if the section is to be included in the CV, or 0 if it should be omitted.
- Column D, Information: The information to be written in the CV file according to section.

### Steps:
Fill Column B: Follow the hierarchical codes as listed in the example. For subsections, continue using the subcodes. If the codes or subcodes are different, the program won't read them. Please try to write the codes in sequential order. Despite this, the program will read them, but unexpected results may arise. Here's the legend:

- 1.0: Name
- 2.1: Residence
- 2.2: e-mail
- 2.3: Cellphone number
- 2.4: Web Page
- 3.0: Professional Statement
- 4.1: University
- 4.2: Degree
- 4.3: Key Accomplishments
- 4.4: Graduation Date
- 4.5: Relevant Coursework
- 5.1: Professional Experience
- 5.2: Role
- 5.3: Location
- 5.4: Starting Date
- 5.5: End Date
- 5.6: Description and Accomplishments
- 6.0: Publications and Presentations
- 7.1: Certifications Group
- 7.2: Certification
- 8.1: Technical Tools
- 8.2: Programming Skills
- 8.3: Languages
- 8.4: Soft Skills
- 8.5: Lab Skills
- 9.0: Honors and Awards
- 10.1: Extracurricular Activities and Leadership Experience
- 10.2: Hobbies
- 11.0: Professional Affiliations
- 12.0: References

Fill Column C: Input 1 if you want this section or subsection to appear in the final CV. Input 0 if this section should be omitted.

Fill Column D: Write the appropiate information for each section and subsection. You can fill this cells even if the corresponding
value on columnc C is zero.

## Notes

Ensure there are no empty rows between sections.
Only input 1 or 0 in the "Included" column. Else, the program won't consider the row or may halt.
For professional experience, if there are multiple roles, repeat the structure for each new role, including all subsections.
The program may read the codes as 1.0 or 1 in the given case.
With hobbies, ensure that only the first one is capitalized.

# GUI

**GUI_01** is a Python application file designed to automate the creation of CV files from an Excel sheet. It uses `ttkbootstrap` to create a modern, styled GUI for interacting with the program. The app reads data from an Excel file, allows users to specify a destination folder and filename, and creates a formatted `.docx` document, with an option to also create a `.pdf` version. 

## Features
- **File and Folder Selection**: Choose an Excel file and a destination folder.
- **Filename Customization**: Specify a name for the output file.
- **Settings**: Options to remove spaces in file names and generate a PDF copy.
- **Instructions**: View detailed instructions within the app for assistance.

## Prerequisites
- Python 3.8+
- `ttkbootstrap` for styling the GUI components.
- `openpyxl`, `docx`, and `tkinter` for file manipulation and GUI handling.

## Setup

1. **Install Required Libraries**:
    ```bash
    pip install ttkbootstrap openpyxl python-docx
    ```
   
2. **Save Required Files**:
   - Save this code in a file, e.g., `GUI_01.py`.
   - Ensure an `Instructions.txt` file is in the same directory, as this will be displayed within the app.

3. **Create Additional Modules**:
   - Implement `fun_write_word` and `fun_read_excel` in a separate file named `Writing_01.py`.
   - `fun_write_word`: This function writes data from the Excel sheet into a `.docx` file.
   - `fun_read_excel`: This function reads data from the specified Excel sheet.

## Usage

1. **Launch the Application**:
   Run the application using:
   ```bash
   python CV_Automator.py


# Writing

This Python script automates the creation of a CV in `.docx` format based on data from an Excel file. It can also convert the `.docx` file to `.pdf` format. This tool provides customizable options for formatting and spacing within the document. 

## Features
- Reads data from a specified Excel file (`.xlsx`).
- Processes the data to generate content for a CV in `.docx` dynamically.
- Allows users to control spacing and conversion to `.pdf`.
- Provides basic error handling for Excel file formatting issues.

## Requirements
- Python 3.x
- `pandas` library for data manipulation.
- `python-docx` library for `.docx` file manipulation.
- `docx2pdf` library for `.pdf` conversion.
- `tkinter` for GUI elements, if needed for additional integration.
- Excel file with predefined formatting for CV data.

### Python Libraries Installation:
To install the required libraries, run:
```bash
pip install pandas python-docx docx2pdf
```

# Functions

This script utilizes the `python-docx` library to create a series of functions aimed at structuring and formatting a CV document in a professional style. Each function is designed to manage a specific section of the CV, with distinct formatting and alignment settings.

## Function Descriptions

### 1. `fun_name`
- **Purpose**: Adds a centered, bold name at the top of the document, followed by a decorative underscore line.
- **Parameters**: Takes in a single list, `f_list_1`, which should contain the name as its first entry.
- **Output**: Formats the name in bold and centers it, adding an underscore line below.

---

### 2. `fun_two`
- **Purpose**: Creates a bulleted list from a string list, where entries are separated by "•".
- **Parameters**: Accepts `f_list_2`, a list of strings for the bulleted items.
- **Output**: Adds each string as a bulleted item in the CV.

---

### 3. `fun_three`
- **Purpose**: Adds a "Personal Statement" section with a heading and justified text.
- **Parameters**: Accepts `f_list_3`, a list of statements.
- **Output**: Adds a heading "Personal Statement" and aligns text in a justified format.

---

### 4. `fun_four`
- **Purpose**: Structures the "Education" section, managing multiple subsections for each entry.
- **Parameters**: Uses `f_list_4` which should include institution names, degrees, and coursework.
- **Output**: Formats each entry with distinct styling for institution, degree, and coursework.

---

### 5. `fun_five`
- **Purpose**: Creates a "Professional Experience" section, detailing job titles, locations, and responsibilities.
- **Parameters**: Takes in `f_list_5`, including job title, location, and description entries.
- **Output**: Aligns text to justify and organizes entries with clear job and location formatting.

---

### 6. `fun_six`
- **Purpose**: Adds a "Publications and Presentations" section.
- **Parameters**: Uses `f_list_6` for entries.
- **Output**: Aligns and formats publication and presentation information.

---

### 7. `fun_seven`
- **Purpose**: Formats a "Certifications" section, displaying certification names with details.
- **Parameters**: Takes `f_list_7`, a list of certification names paired with details.
- **Output**: Aligns each certification entry.

---

### 8. `fun_eight`
- **Purpose**: Configures a custom section using nested flags and strings, for advanced formatting.
- **Parameters**: Accepts `f_list_8`.
- **Output**: Handles complex data, potentially involving multi-level formatting.

---

### 9. fun_nine
- **Purpose**: Adds an "Honors and Awards" section to the document, listing awards in a centered, bold header followed by left-aligned entries.
- **Parameters**: Takes `f_list_nine`, a list containing the honors and awards data, where each entry is a list that includes the award description.
- **Output**: The function prints each award entry to the console while formatting it as left-aligned text in the document.

---

### 10. fun_ten
- **Purpose: Adds an "Extracurricular Activities" section with two parts: a list of activities and a separate "Hobbies" list. Activities are categorized and formatted accordingly.
- **Parameters: Takes `f_list_ten`, a list containing extracurricular activities, where each entry includes a flag for classification and a description.
- **Output**: The function prints each activity and hobby to the console and formats them into the document with specific alignment and boldness for headers.

---

### 11. fun_eleven
- **Purpose**: Creates a "References" section in the document, with a centered, bold heading followed by left-aligned entries for each reference.
- **Parameters**: Takes `f_list_eleven`, a list containing references, where each entry holds a reference description.
- **Output**: The function prints each reference to the console and formats the reference details as left-aligned text under the "References" heading.

---

### 12. fun_twelve
- **Purpose**: Adds a "Professional Affiliations" section to the document, with a bold heading and left-aligned entries for each affiliation.
- **Parameters**: Takes `f_list_twelve`, a list containing professional affiliations, where each entry holds the name of an affiliation.
- **Output**: The function prints each affiliation to the console and formats it as left-aligned text under the "Professional Affiliations" heading.

---

## Formatting Standards
The alignment and font styles are consistently set to maintain a clean and professional appearance. This structure provides a modular approach, allowing easy modifications or additions to each section as needed.

## Prerequisites
- Python 3.8+
- `docx` for file manipulation.

## Setup

1. **Install Required Libraries**:
    ```bash
    pip install python-docx
    ```
