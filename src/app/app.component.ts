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
      pageTitle: 'List of Events',
      colsWidth: [10, 15, 35, 35],
      tables: [
        {
          rowsBefore: [
            { rows: [{ col: 'Dummy1', offset: 2 }, { col: 'Dummy2' }], addBlankRow: true },
            { rows: [{ col: 'Dummy1', offset: 2 }, { col: 'Dummy2' }] },
          ],
          tableColumns: [
            { key: 'sNo', label: 'S No.' },
            { key: 'eventName', label: 'Event Name' },
            { key: 'eventStartDate', label: 'Start Date' },
            { key: 'eventEndDate', label: 'End Date' },
          ],
          tableData: [
            {
              sNo: 1,
              eventName: 'Event 1',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 2,
              eventName: 'Event 2',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 3,
              eventName: 'Event 3',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 4,
              eventName: 'Event 4',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 5,
              eventName: 'Event 5',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 6,
              eventName: 'Event 6',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 7,
              eventName: 'Event 7',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 8,
              eventName: 'Event 8',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 9,
              eventName: 'Event 9',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 10,
              eventName: 'Event 10',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 11,
              eventName: 'Event 11',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
          ],
        },
        {
          tableColumns: [
            { key: 'sNo', label: 'S No.' },
            { key: 'eventName', label: 'Event Name' },
            { key: 'eventStartDate', label: 'Start Date' },
            { key: 'eventEndDate', label: 'End Date' },
          ],
          tableData: [
            {
              sNo: 1,
              eventName: 'Event 1',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 2,
              eventName: 'Event 2',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 3,
              eventName: 'Event 3',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 4,
              eventName: 'Event 4',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 5,
              eventName: 'Event 5',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 6,
              eventName: 'Event 6',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 7,
              eventName: 'Event 7',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 8,
              eventName: 'Event 8',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 9,
              eventName: 'Event 9',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 10,
              eventName: 'Event 10',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
            {
              sNo: 11,
              eventName: 'Event 11',
              eventStartDate: '2021-08-11T12:46:12.354Z',
              eventEndDate: '2021-08-15T12:46:12.354Z',
            },
          ],
        },
      ],
    });
  }
}
