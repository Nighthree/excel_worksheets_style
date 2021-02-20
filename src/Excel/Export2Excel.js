//Export2Excel.js
/* eslint-disable */
require('script-loader!file-saver');
require('./Blob.js');//转二进制用  这边要写你的blob的实际地址
require('script-loader!xlsx/dist/xlsx.core.min');
import XLSX from 'xlsx-style'

function generateArray(table) {
  var out = [];
  var rows = table.querySelectorAll('tr');
  var ranges = [];
  for (var R = 0; R < rows.length; ++R) {
    var outRow = [];
    var row = rows[R];
    var columns = row.querySelectorAll('td');
    for (var C = 0; C < columns.length; ++C) {
      var cell = columns[C];
      var colspan = cell.getAttribute('colspan');
      var rowspan = cell.getAttribute('rowspan');
      var cellValue = cell.innerText;
      if (cellValue !== "" && cellValue == +cellValue) cellValue = +cellValue;

      //Skip ranges
      ranges.forEach(function (range) {
        if (R >= range.s.r && R <= range.e.r && outRow.length >= range.s.c && outRow.length <= range.e.c) {
          for (var i = 0; i <= range.e.c - range.s.c; ++i) outRow.push(null);
        }
      });

      //Handle Row Span
      if (rowspan || colspan) {
        rowspan = rowspan || 1;
        colspan = colspan || 1;
        ranges.push({
          s: {
            r: R,
            c: outRow.length
          },
          e: {
            r: R + rowspan - 1,
            c: outRow.length + colspan - 1
          }
        });
      };

      //Handle Value
      outRow.push(cellValue !== "" ? cellValue : null);

      //Handle Colspan
      if (colspan)
        for (var k = 0; k < colspan - 1; ++k) outRow.push(null);
    }
    out.push(outRow);
  }
  return [out, ranges];
};

function datenum(v, date1904) {
  if (date1904) v += 1462;
  var epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data, opts) {
  console.log(data);
  var ws = {};
  var range = {
    s: {
      c: 10000000,
      r: 10000000
    },
    e: {
      c: 0,
      r: 0
    }
  };
  for (var R = 0; R != data.length; ++R) {
    for (var C = 0; C != data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R;
      if (range.s.c > C) range.s.c = C;
      if (range.e.r < R) range.e.r = R;
      if (range.e.c < C) range.e.c = C;
      // console.log(`data[${R}][${C}]`, data[R][C])
      // var cell = {
      //   ...data[R][C]
      // };
      var cell = { v: data[R][C] };
      if (cell.v == null) continue;
      var cell_ref = XLSX.utils.encode_cell({
        c: C,
        r: R
      });

      if (typeof cell.v === 'number') cell.t = 'n';
      else if (typeof cell.v === 'boolean') cell.t = 'b';
      else if (cell.v instanceof Date) {
        cell.t = 'n';
        cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      } else cell.t = 's';

      ws[cell_ref] = cell;
      // console.log(`ws[${cell_ref}]`,ws[cell_ref]);
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  console.log('ws', ws);
  // ws['!merges'] = [
  //   // s：开始start
  //   // e：结束end
  //   // r：行row
  //   // c: 列columns
  //   { s: { r: 0, c: 0 }, e: { r: 0, c: data.length } }, // 第一行，第一列到第三列合并
  //   // {s: {r: 0, c: 3}, e: {r: 0, c: 4}}, // 第一行，第三列到第四列合并
  // ]
  return ws;
}

function Workbook() {
  if (!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

export function export_table_to_excel(id) {
  var theTable = document.getElementById(id);
  console.log('a')
  var oo = generateArray(theTable);
  var ranges = oo[1];

  /* original data */
  var data = oo[0];
  var ws_name = "SheetJS";
  console.log(data);

  var wb = new Workbook(),
    ws = sheet_from_array_of_arrays(data);

  /* add ranges to worksheet */
  // ws['!cols'] = ['apple', 'banan'];
  ws['!merges'] = ranges;

  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;

  var wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'
  });

  saveAs(new Blob([s2ab(wbout)], {
    type: "application/octet-stream"
  }), "test.xlsx")
}

function formatJson(jsonData) {
  console.log(jsonData)
}
export function export_json_to_excel(th, jsonData, defaultTitle, sheetArray) {
  console.log(th);
  console.log(jsonData);
  /* original data */
  // 產生一個的分頁
  // var data = jsonData;

  // data.unshift(th);

  // var ws_name = "SheetJS";

  // var wb = new Workbook(),
  //   ws = sheet_from_array_of_arrays(data);

  // /* add worksheet to workbook */
  // wb.SheetNames.push(ws_name);
  // wb.Sheets[ws_name] = ws;

  // 產生多個分頁
  // 重新組範例的不適合
  var data = [];
  for (let i = 0; i < jsonData.length; i++) {
    let arr = [];
    arr.push(th[i]);
    for (let j = 0; j < jsonData[i].length; j++) {
      arr.push(jsonData[i][j])
    }
    console.log('arr', arr);
    data.push(arr);
  }
  // console.log('進來的 data', data);
  //添加标题
  // for (var i = 0; i < th.length; i++) {
  //   data[i].unshift(th[i]);
  //   console.log(`th[${i}]`,th[i]);
  //   console.log(`data[${i}].unshift(th[${i}])`, data[i].unshift(th[i]));
  //   console.log(`data[${i}]`, data[i]);
  // }

  //这里是定义sheet的名称
  var ws_name = sheetArray;

  var wb = new Workbook(), ws = [];
  //数据转换
  for (var j = 0; j < th.length; j++) {
    ws.push(sheet_from_array_of_arrays(data[j]))
  }

  /* add worksheet to workbook */
  //生成多个sheet
  for (var k = 0; k < th.length; k++) {
    wb.SheetNames.push(ws_name[k])
    wb.Sheets[ws_name[k]] = ws[k]
  }

  var wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    bookSST: false,
    autoWidth: true,
    type: 'binary'
  });
  var title = defaultTitle || '列表'

  saveAs(new Blob([s2ab(wbout)], {
    type: "application/octet-stream"
  }), title + ".xlsx")
};