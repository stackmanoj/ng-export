import { Component, OnInit } from '@angular/core';
import { ExportExcelService } from './services/export-excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'sample-ng-app';
  constructor(private excelService: ExportExcelService) {

  }
  ngOnInit(): void {

  }
  export(type) {
    let json = [{ UserName: 'test name', Email: 'mk@g.com' }, { UserName: 'test name 2', Email: 'mk2@g.com' }]
    this.excelService.export(type, json, 'Users', ['User Name', 'Email'], ['UserName', 'Email']);
  }
}
