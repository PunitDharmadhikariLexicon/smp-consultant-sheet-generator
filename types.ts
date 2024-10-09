import { CellValue } from "exceljs";

export type Consultant = {
  fullName: CellValue;
  email: CellValue;
  jobTitle: CellValue;
  grade: CellValue;
  discipline: CellValue;
  location: CellValue;
};

export type Skill = {
  name: string;
  category: string;
};

export type Row = {
  col: number;
  value: string;
};

export type InputFile = {
  consultants: ConsultantData;
  // Add Skills and Categories when available
};

type ConsultantData = {
  name: string;
  sheet: string;
};
