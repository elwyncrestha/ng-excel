import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { ExcelPreference } from './excel-preference.model';

@Injectable({
  providedIn: 'root',
})
export class ExcelService {
  private readonly fileExtension = '.xlsx';
  private readonly defaultPreferences: ExcelPreference = {
    fileName: 'excel',
    sheetName: 'sheet'
  };

  private XLSX_Utils: XLSX.XLSX$Utils;
  private XLSX_WriteFile: (data: XLSX.WorkBook, filename: string, opts?: XLSX.WritingOptions) => any;
  private workSheet!: XLSX.WorkSheet;

  constructor() {
    this.XLSX_Utils = XLSX.utils;
    this.XLSX_WriteFile = XLSX.writeFile;
  }

  async generate(preferences: ExcelPreference): Promise<void> {
    this.export(preferences);
  }

  private export(preferences: ExcelPreference): void {
    this.initWorkSheet();
    this.exportToFile(preferences);
  }

  private initWorkSheet(): void {
    this.workSheet = this.XLSX_Utils.json_to_sheet([]);
  }

  private exportToFile(preferences: ExcelPreference): void {
    /* generate workSheet and add the worksheet */
    const workbook = this.XLSX_Utils.book_new();
    this.XLSX_Utils.book_append_sheet(workbook, this.workSheet, preferences?.sheetName ?? this.defaultPreferences.sheetName);
    /* save to file */
    this.XLSX_WriteFile(workbook, `${preferences?.fileName ?? this.defaultPreferences.fileName}${this.fileExtension}`);
  }
}
