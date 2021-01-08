import { Component, OnInit } from '@angular/core';
import * as  Excel from 'exceljs/dist/exceljs.min.js';

@Component({
  selector: 'app-exceltojson',
  templateUrl: './exceltojson.component.html',
  styleUrls: ['./exceltojson.component.css']
})
export class ExceltojsonComponent implements OnInit {
  ExcelUploadData: object[] = [];

  constructor() { }

  ngOnInit(): void {
  }

  readExcel(event) {
    let ExcelColumnHeaders: string[] = [];
    const ExcelValuesArray: object[] = [];
    // debugger;
    const workbook = new Excel.Workbook();
    const target: DataTransfer = (event.target) as DataTransfer;
    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
      alert('only Single Excel File Upload will processed');
    }

    const arryBuffer = new Response(target.files[0]).arrayBuffer();
    arryBuffer.then((data) => {
      workbook.xlsx.load(data).then(() => {
        console.log(workbook);
        const worksheet = workbook.getWorksheet(1);
        console.log('rowCount:', worksheet.rowCount);
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) {
            ExcelColumnHeaders = row.values.slice(1);
          } else if (rowNumber > 1) {
            const obj: object = new Object();
            const columnsvalues = row.values.slice(1);
            ExcelColumnHeaders.forEach((element, index) => {
              obj[element] = columnsvalues[index];
            });
            ExcelValuesArray.push(obj);
            console.log(ExcelValuesArray);

          }
          console.table('Row: ' + rowNumber + ' Value: ' + row.values);

        });
      });
    });
    this.ExcelUploadData = ExcelValuesArray;
  }

}
