import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import { Workbook } from 'exceljs';
import * as jsPDF from 'jspdf';
import 'jspdf-autotable';

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Injectable()
export class ExportExcelService {

    constructor() { }

    public export(type: 'pdf' | 'excel', json: any[], reportHeader: string, columnNames: string[], dataProperties: string[]) {
        if (type == 'pdf') {
            this.exportAsPdfFile(json, reportHeader, columnNames, dataProperties);
        }
        else if (type == 'excel') {
            this.exportAsExcelFile(json, reportHeader, columnNames, dataProperties);
        }
    }

    public exportAsPdfFile(json: any[], reportHeader: string, columnNames: string[], dataProperties: string[]): void {
        json = this.generateJson(json, columnNames, dataProperties);
        var doc = new jsPDF();
        var col = columnNames;
        var rows = [];
        json.forEach(element => {
            rows.push(element);
        });
        doc.autoTable(col, rows, {
            theme: 'grid',
            styles: { fillColor: '#fff' },
            headerStyles: { fillColor: 'brown', textColor: 'black' },
            beforePageContent: function (data) {
                doc.text(reportHeader, 20, 10);
            },
            margin: [20, 10, 20, 10]
        });
        doc.save(`${reportHeader}.pdf`);
    }
    public exportAsExcelFile(json: any[], reportHeader: string, columnNames: string[], dataProperties: string[]): void {
        json = this.generateJson(json, columnNames, dataProperties);
        let workbook = new Workbook();
        let worksheet = workbook.addWorksheet('Users');
        let titleRow = worksheet.addRow([reportHeader]);
        titleRow.font = { name: 'Calibri', family: 4, size: 16, bold: true };
        titleRow.alignment = { horizontal: 'center' };
        worksheet.mergeCells(`A1:${String.fromCharCode(64 + columnNames.length)}1`);

        //Add Header Row
        let headerRow = worksheet.addRow(columnNames);

        // Cell Style : Fill and Border
        headerRow.eachCell((cell, number) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFF00' },
                bgColor: { argb: 'FF0000FF' },
            }
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
        })

        json.forEach(d => {
            let row = worksheet.addRow(d);
        });
        //Generate Excel File with given name
        workbook.xlsx.writeBuffer().then((data) => {
            let blob = new Blob([data], { type: EXCEL_TYPE });
            this.saveAsExcelFile(blob, reportHeader);
        })

    }

    private generateJson(json: any[], columnNames: string[], dataProperties: string[]) {
        let newJson = [];
        for (let j = 0; j < json.length; j++) {
            let obj: any = [];
            for (let i = 0; i < dataProperties.length; i++) {
                obj.push(json[j][dataProperties[i]]);
            }
            newJson.push(obj);
        }
        return newJson;
    }

    private saveAsExcelFile(buffer: any, fileName: string): void {
        const data: Blob = new Blob([buffer], {
            type: EXCEL_TYPE
        });
        FileSaver.saveAs(data, fileName + '_export_' + new Date().getTime() + EXCEL_EXTENSION);
    }

}