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
import { ExcelPreference } from './excel.model';

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
    this.renderTable();
    this.export();
  }

  private initWorkSheet(): void {
    const { pageTitle } = this.preferences;
    this.workSheet = pageTitle
      ? this.XLSX_Utils.json_to_sheet([{ A: pageTitle }], {
          header: ['A'],
          skipHeader: true,
        })
      : this.XLSX_Utils.json_to_sheet([]);

    if (pageTitle) {
      this.workSheet['!rows'] = [{ hpt: 25 }];
      this.addBlankRow();
    }
  }

  private renderTable(): void {
    const { columns } = this.preferences;
    if (!columns) {
      return;
    }
    // render columns
    this.XLSX_Utils.sheet_add_aoa(this.workSheet, [columns], { origin: -1 });
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
