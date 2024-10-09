import ExcelJS from "exceljs";
import { Consultant, InputFile } from "./types";

const inputFile: InputFile = {
  consultants: {
    name: "Skills matrix.xlsx",
    sheet: "People Data",
  },
  // Add Skills and Categories file when available
};

const outputFilename = "Endava-Skills-Matrix.xlsx";

(async () => {
  const consultants = await loadConsultants(inputFile.consultants);

  const skillsAndCategories: Record<string, string[]> = {
    Category1: ["skill1", "skill2"],
    Category2: ["skill3", "skill4"],
  };

  writeSkillsMatrix(outputFilename, consultants, skillsAndCategories);
})();

// async function loadSkillsAndCategories(
//   filename: string,
//   sheetName: string
// ): Promise<Record<string, Record<string, string[]>>> {
//   await workbook.xlsx.readFile(filename);
//   const sheet = workbook.getWorksheet(sheetName);

//   const consultants: Consultant[] = [];

//   if (sheet) {
//     sheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
//       if (rowNumber > 1) {
//         consultants.push({
//           fullName: row.getCell(1).value,
//           email: row.getCell(6).value,
//           jobTitle: row.getCell(7).value,
//           grade: row.getCell(8).value,
//           discipline: row.getCell(9).value,
//           location: row.getCell(11).value,
//         });
//       }
//     });
//   }

//   return consultants;
// }

const headerRowFont = {
  size: 16,
  bold: true,
};

const rowFont = {
  size: 12,
};

const consultantNameFont = {
  size: 14,
};

const alignment: Partial<ExcelJS.Alignment> = {
  horizontal: "center",
  vertical: "middle",
};

function writeSkillsMatrix(
  filename: string,
  consultants: Consultant[],
  skillsAndCategories: Record<string, string[]>
) {
  const workbook = new ExcelJS.Workbook();
  const sheet1 = workbook.addWorksheet("Consultants");

  sheet1.columns = [
    { header: "Full Name", key: "fullName", width: 30 },
    { header: "Email", key: "email", width: 30 },
    { header: "Location", key: "location", width: 20 },
    { header: "Job Title", key: "jobTitle", width: 20 },
    { header: "Discipline", key: "discipline", width: 20 },
    { header: "Grade", key: "grade", width: 20 },
    { header: "Profile Link", key: "link", width: 20 },
  ];

  const headerRow = sheet1.getRow(1);
  headerRow.font = headerRowFont;

  sheet1.autoFilter = {
    from: "A1",
    to: "F1",
  };

  const sheet2 = workbook.addWorksheet("Skill Proficiencies");

  let currentColIndex = 2;
  const categoryRow = sheet2.getRow(1);
  const skillRow = sheet2.getRow(2);
  categoryRow.font = headerRowFont;

  categoryRow.alignment = alignment;
  skillRow.font = headerRowFont;

  const consultantColumn = sheet2.getColumn("A");
  consultantColumn.width = 30;

  Object.keys(skillsAndCategories).forEach((category) => {
    const skills = skillsAndCategories[category];
    const startCol = currentColIndex;
    const endCol = currentColIndex + skills.length - 1;
    sheet2.mergeCells(1, startCol, 1, endCol);
    categoryRow.getCell(startCol).value = category;

    skills.forEach((skill, index) => {
      skillRow.getCell(startCol + index).value = skill;
    });

    currentColIndex += skills.length;
  });

  sheet2.getCell("A2").value = "Name";

  consultants.forEach((consultant, consultantIndex) => {
    const rowIndex = consultantIndex + 3;
    const consultantCell = sheet2.getCell(`A${rowIndex}`);
    consultantCell.value = consultant.fullName;
    consultantCell.font = consultantNameFont;

    const row = sheet1.addRow(consultant);
    row.font = rowFont;

    const linkCell = sheet1.getCell(`G${consultantIndex + 2}`);
    linkCell.value = {
      text: `Go to ${consultant.fullName}'s Skills`,
      hyperlink: `#'Skill Proficiencies'!A${rowIndex}`,
    };
    linkCell.font = { color: { argb: "FF0000FF" }, underline: true };

    let colOffset = 0;

    Object.keys(skillsAndCategories).forEach((category) => {
      const skills = skillsAndCategories[category];
      skills.forEach((_, skillIndex) => {
        const colIndex = 2 + colOffset + skillIndex;
        const cell = sheet2.getCell(rowIndex, colIndex);
        cell.dataValidation = {
          type: "list",
          allowBlank: true,
          formulae: ['"1,2,3,4,5"'],
          showErrorMessage: true,
          errorTitle: "Invalid Proficiency Level",
          error: "Please select a value from 1 to 5 or leave blank.",
          promptTitle: "Proficiency Level",
          prompt: "Select a value between 1 and 5.",
        };
      });
      colOffset += skills.length;
    });
  });

  workbook.xlsx
    .writeFile(filename)
    .then(() => {
      console.log(`Excel file '${filename}' created successfully.`);
    })
    .catch((err: unknown) => {
      console.log("Error creating Excel file:", err);
    });
}

async function loadConsultants(
  consultantData: typeof inputFile.consultants
): Promise<Consultant[]> {
  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.readFile(consultantData.name);
  const sheet = workbook.getWorksheet(consultantData.sheet);

  const consultants: Consultant[] = [];

  if (sheet) {
    sheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
      if (rowNumber > 1) {
        consultants.push({
          fullName: row.getCell(1).value,
          email: row.getCell(6).value,
          jobTitle: row.getCell(7).value,
          grade: row.getCell(8).value,
          discipline: row.getCell(9).value,
          location: row.getCell(11).value,
        });
      }
    });
  }

  return consultants;
}
