export interface ExcelPreference {
  fileName?: string;
  sheetName?: string;
  pageTitle?: string;
  // col attributes
  colsWidth?: number[];
  // table attributes
  tables?: {
    rowsBefore?: CustomTableRow[];
    tableColumns: { key: string; label: string }[];
    tableData: { [key: string]: any }[];
    rowsAfter?: CustomTableRow[];
  }[];
}

export interface CustomTableRow {
  rows: {
    col: any;
    offset?: number;
  }[];
  addBlankRow?: boolean;
}
