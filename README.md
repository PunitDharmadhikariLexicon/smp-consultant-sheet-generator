# Endava Skills Matrix Generator

This tool generates a blank Excel file based on consultant and skill data. It extracts consultant information from one file and skill categories from another, then organizes the data into two sheets:

1. **Consultants Sheet**:

   - Contains consultant information: Full Name, Email, Location, Job Title, Discipline, Grade, and a hyperlink to their skills in the second sheet.

2. **Skill Proficiencies Sheet**:
   - Lists the skills categorized into predefined groups. Each consultant is listed along with dropdown options (1-5) for their skill proficiency ratings.

## Prerequisites

Before running the tool, ensure you have the following:

- **Node.js** installed.
- **ExcelJS** package installed: `npm install exceljs`.

## Setup

1. Clone the repository.
2. You need to manually add the following files to the root directory:
   - **Skills matrix.xlsx**: This file should contain the consultant data under the sheet named `People Data`.
   - **skills 4.xlsx**: This file should contain the skills and their categories under the sheet named `Result 1`.
3. Run `npm i` in the repo.

These files are not committed to the repository for confidentiality reasons.

## Running the Tool

Once the required files are added, simply run the tool using:

```bash
npm run sheet
```

This will generate the output Excel file in the root directory.

## Output

Consultants Sheet: Contains consultant details and links to their corresponding skill row in the Skill Proficiencies Sheet.
Skill Proficiencies Sheet: Lists skills and allows selecting proficiency levels (1-5) for each consultant.
