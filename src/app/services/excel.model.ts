export interface ExcelPreference {
  fileName?: string;
  sheetName?: string;
  pageTitle?: string;
  // col attributes
  colsWidth?: number[];
  // table attributes
  tableColumns?: { key: string; label: string; }[];
  tableData?: { [key: string]: any; }[];
}
