import { Injectable } from '@angular/core';
import {
  CellObject,
  utils,
  WorkBook,
  WorkSheet,
  writeFile,
  WritingOptions,
  XLSX$Utils,
} from 'xlsx';
import { CustomTableRow, ExcelPreference } from './excel.model';

@Injectable({
  providedIn: 'root',
})
export class ExcelService {
  private readonly fileExtension = '.xlsx';
  private readonly defaultPreferences: ExcelPreference = {
    fileName: 'excel',
    sheetName: 'sheet',
  };

  private XLSX_Utils: XLSX$Utils;
  private XLSX_WriteFile: (
    data: WorkBook,
    filename: string,
    opts?: WritingOptions
  ) => any;
  private readonly addBlankRow = () =>
    this.XLSX_Utils.sheet_add_aoa(this.workSheet, [[]], { origin: -1 });

  private workSheet!: WorkSheet;
  private preferences!: ExcelPreference;

  constructor() {
    this.XLSX_Utils = utils;
    this.XLSX_WriteFile = writeFile;
  }

  async generate(preferences: ExcelPreference): Promise<void> {
    this.preferences = preferences;
    this.initWorkSheet();
    this.addPageTitle();
    this.renderTables();
    this.setColsWidth();
    this.export();
  }

  private initWorkSheet(): void {
    this.workSheet = {};
  }

  private addPageTitle(): void {
    const { pageTitle } = this.preferences;
    if (!pageTitle) {
      return;
    }

    this.XLSX_Utils.sheet_add_json(this.workSheet, [{ A: pageTitle }], {
      header: ['A'],
      skipHeader: true,
    });
    this.workSheet['!rows'] = [{ hpt: 25 }];
    this.addBlankRow();
  }

  private renderTables(): void {
    const { tables } = this.preferences;
    tables?.forEach((table) => {
      const { rowsBefore, tableColumns, tableData, rowsAfter } = table;
      if (!tableColumns) {
        return;
      }
      // add rows before table if any
      if (rowsBefore) {
        this.addRows(rowsBefore);
      }
      // render columns
      this.XLSX_Utils.sheet_add_aoa(
        this.workSheet,
        [tableColumns.map((c) => c.label)],
        { origin: -1 }
      );
      // render data
      const rows = new Array<Array<any>>();
      tableData?.forEach((data) => {
        const row = new Array<any>();
        tableColumns
          .map((c) => c.key)
          .forEach((c) => {
            row.push(data[c] ?? '');
          });
        rows.push(row);
      });
      this.XLSX_Utils.sheet_add_aoa(this.workSheet, rows, { origin: -1 });
      this.addBlankRow();
      // add rows after table if any
      if (rowsAfter) {
        this.addRows(rowsAfter);
      }
    });
  }

  private addRows(addingRows: CustomTableRow[]) {
    addingRows.forEach((add) => {
      const newRow: any[] = [];
      add.rows.forEach((rowData) => {
        newRow.push(rowData.col);
        if (rowData.offset) {
          for (let i = 0; i < rowData.offset; i++) {
            newRow.push('');
          }
        }
      });
      this.XLSX_Utils.sheet_add_aoa(this.workSheet, [newRow], { origin: -1 });
      if (add.addBlankRow) {
        this.addBlankRow();
      }
    });
  }

  private setColsWidth(): void {
    const { colsWidth } = this.preferences;
    if (!colsWidth || colsWidth.length < 1) {
      return;
    }

    if (!this.workSheet['!cols']) {
      this.workSheet['!cols'] = [];
    }
    colsWidth.forEach((w, i) => {
      if (this.workSheet['!cols']) {
        this.workSheet['!cols'][i] = { width: w };
      }
    });
  }

  private export(): void {
    const {
      fileName = this.defaultPreferences.fileName,
      sheetName = this.defaultPreferences.sheetName,
    } = this.preferences;
    /* generate workSheet and add the worksheet */
    const workbook = this.XLSX_Utils.book_new();
    this.XLSX_Utils.book_append_sheet(workbook, this.workSheet, sheetName);
    /* save to file */
    this.XLSX_WriteFile(workbook, `${fileName}${this.fileExtension}`);
  }
}
