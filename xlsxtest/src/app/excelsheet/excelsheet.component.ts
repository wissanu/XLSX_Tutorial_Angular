import { Component, OnInit } from '@angular/core';
import { Data } from '@angular/router';
import * as XLSX from 'xlsx';
import { ExcelService } from './../excel.service';

@Component({
  selector: 'app-excelsheet',
  templateUrl: './excelsheet.component.html',
  styleUrls: ['./excelsheet.component.css'],
})
export class ExcelsheetComponent implements OnInit {
  data: any;
  readdata: any;
  constructor(private excelService: ExcelService) {
    this.data = [
      [2019, 1, '50', '20', '25', '20'],
      [2019, 2, '80', '20', '25', '20'],
      [2019, 3, '120', '20', '25', '20'],
      [2019, 4, '75', '20', '25', '20'],
      [2019, 5, '60', '20', '25', '20'],
      [2019, 6, '80', '20', '25', '20'],
      [2019, 7, '95', '20', '25', '20'],
      [2019, 8, '55', '20', '25', '20'],
      [2019, 9, '45', '20', '25', '20'],
      [2019, 10, '80', '20', '25', '20'],
      [2019, 11, '90', '20', '25', '20'],
      [2019, 12, '110', '20', '25', '20'],
    ];
  }

  ngOnInit(): void {}

  onFileChange(evt: any) {
    const target: DataTransfer = <DataTransfer>evt.target;

    if (target.files.length !== 1) throw Error('not support multiple files.');

    const reader: FileReader = new FileReader();

    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];
      console.log(ws);
      this.readdata = XLSX.utils.sheet_to_json(ws, { header: 1 });
      console.log(this.readdata);
    };

    reader.readAsBinaryString(target.files[0]);
  }

  generateExcel() {
    // console.log('called');
    this.excelService.generateExcel();
  }
}
