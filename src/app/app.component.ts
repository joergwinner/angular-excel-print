import { Component, OnInit } from '@angular/core';

import { utils, write, WorkBook } from 'xlsx';

import { saveAs } from 'file-saver';


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'app';

  PERSONS: any[] = [
    {
        TransactionType:1,
        Username: 'Thomas' ,
        TransactionTime: '13:15',
        ImageSize: 21,
        IP: '117.247.92.158',
        Project:'Project 1',
        ImageName:'Name 1'
    },
    {
        TransactionType: 2,
        Username: 'Adam',
        TransactionTime: '14:15',
        ImageSize: 12,
        IP: '548.247.92.158',
        Project:'Project 2',
        ImageName:'Name 2'
    },
    {
        TransactionType: 3,
        Username: 'Steve',
        TransactionTime: '10:45',
        ImageSize: 38,
        IP: '117.888.92.124',
        Project:'Project 3',
        ImageName:'Name 3'
    },
    {
        TransactionType: 4,
        Username: 'Adam',
        TransactionTime: '17:25',
        ImageSize: 38,
        IP: '989.247.52.158',
        Project:'Project 4',
        ImageName:'Name 4'
    },
    {
        TransactionType: 5,
        Username: 'Col',
        TransactionTime: '17:00',
        ImageSize: 38,
        IP: '658.245.87.987',
        Project:'Project 5',
        ImageName:'Name 5'
    },
    {
        TransactionType: 6,
        Username: 'Mac',
        TransactionTime: '21:10',
        ImageSize: 38,
        IP: '254.254.256.254',
        Project:'Project 6',
        ImageName:'Name 6'
    },
    {
        TransactionType: 7,
        Username: 'Bruse',
        TransactionTime: '22:21',
        ImageSize: 38,
        IP: '888.666.11.235',
        Project:'Project 7',
        ImageName:'Name 7'
    },
    {
        TransactionType: 8,
        Username: 'Luke',
        TransactionTime: '23:11',
        ImageSize: 38,
        IP: '879.852.65.879',
        Project:'Project 8',
        ImageName:'Name 8'
    },
    {
        TransactionType: 9,
        Username: 'Lucky',
        TransactionTime: '19:33',
        ImageSize: 38,
        IP: '458.785.54.236',
        Project:'Project 9',
        ImageName:'Name 9'
    },
    {
        TransactionType: 10,
        Username: 'John',
        TransactionTime: '14:20',
        ImageSize: 38,
        IP: '874.569.21.158',
        Project:'Project 10',
        ImageName:'Name 10'
    },
    {
        TransactionType: 11,
        Username: 'Rex',
        TransactionTime: '10:52',
        ImageSize: 38,
        IP: '117.247.92.158',
        Project:'Project 11',
        ImageName:'Name 11'
    },
    {
        TransactionType: 12,
        Username: 'Steve',
        TransactionTime: '8:20',
        ImageSize: 38,
        IP: '548.247.54.158',
        Project:'Project 12',
        ImageName:'Name 12'
    },
    {
        TransactionType: 13,
        Username: 'Aka',
        TransactionTime: '9:33',
        ImageSize: 38,
        IP: '458.785.54.236',
        Project:'Project 13',
        ImageName:'Name 13'
    }
];

  ngOnInit() {

  }
  print() {
    let printContents, popupWin, htmlToPrint;
    printContents = document.getElementById('print');
    htmlToPrint = '' +
        '<style type="text/css">'+
        'table, td, th {'+
           'border: 1px solid grey;'+
            'text-align: left;'+
            'font-family: Segoe UI;'+
        '}'+
        'table {'+
            'border-collapse: collapse;'+
            'width: 100%;'+
       '}'+
        'th, td {'+
            'padding: 6px;'+
            '}' +
            '</style>';
            
    htmlToPrint += printContents.outerHTML;
    popupWin = window.open('', '_blank', 'top=0,left=200,height=1000,width=1000');
    popupWin.document.open();
    popupWin.document.write(`
    <html>
    <body onload="window.print();window.close()">${htmlToPrint}</body>
    </html>`);
    popupWin.document.close();
  }

  super(){
    const ws_name = 'SuperOSM';
    const wb: WorkBook = { SheetNames: [], Sheets: {} };
    const ws: any = utils.json_to_sheet(this.PERSONS);
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;
    const wbout = write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });

    function s2ab(s) {
      const buf = new ArrayBuffer(s.length);
      const view = new Uint8Array(buf);
      for (let i = 0; i !== s.length; ++i) {
        view[i] = s.charCodeAt(i) & 0xFF;
      };
      return buf;
    }

    saveAs(new Blob([s2ab(wbout)], { type: 'application/octet-stream' }), 'FSPKO8.xlsx');
  }
}

