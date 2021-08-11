import { Component } from '@angular/core';
import { ExcelService } from './services/excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  title = 'ng-excel';

  constructor(private readonly excelService: ExcelService) {}

  generate(): void {
    this.excelService.generate({
      fileName: 'ng-excel',
      sheetName: 'Default',
      // pageTitle: 'List of Events',
      columns: ['S No.', 'Event Name', 'Start Date', 'End Date']
    });
  }
}
