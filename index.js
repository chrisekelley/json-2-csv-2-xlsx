var converter = require('json-2-csv');
var fs = require('fs');
var XLSX = require('xlsx');

var ws_name = "SheetJS";

function Workbook() {
  if(!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

var wb = new Workbook()

var filename = 'A.csv';

function datenum(v, date1904) {
  if(date1904) v+=1462;
  var epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function create_array_of_arrays(data) {
  var doc = [];
  var arr = data.split('\n');
  for(var i = 0; i != arr.length; ++i) {
    lineArr = arr[i].split(",")
    doc[i] = lineArr;
  }
  //Delete the last empty array
  if (doc[arr.length-1] == "") {
    var deleted = doc.splice((arr.length-1), 1);
  }
  return doc;
}

function sheet_from_array_of_arrays(data, opts) {
  var ws = {};
  var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
  for(var R = 0; R != data.length; ++R) {
    for(var C = 0; C != data[R].length; ++C) {
      if(range.s.r > R) range.s.r = R;
      if(range.s.c > C) range.s.c = C;
      if(range.e.r < R) range.e.r = R;
      if(range.e.c < C) range.e.c = C;
      var cell = {v: data[R][C] };
      if(cell.v == null) continue;
      var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

      if(typeof cell.v === 'number') cell.t = 'n';
      else if(typeof cell.v === 'boolean') cell.t = 'b';
      else if(cell.v instanceof Date) {
        cell.t = 'n'; cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      }
      else cell.t = 's';

      ws[cell_ref] = cell;
    }
  }
  if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
}

var json2csvCallback = function (err, data) {
  if (err) throw err;
  //console.log("data: " + JSON.stringify(data));
  //fs.writeFile(filename, data, function(err) {
  //  if(err) {
  //    console.log(err);
  //  } else {
      ws_name = 'A';
      if (filename == 'B.csv') {
        ws_name = 'B';
      }
      console.log("Creating worksheet: " + ws_name);
      /* add worksheet to workbook */
      wb.SheetNames.push(ws_name);
      var arr = create_array_of_arrays(data)
      var ws = sheet_from_array_of_arrays(arr);
      wb.Sheets[ws_name] = ws;

      if (filename == 'B.csv') {
        console.log("Writing File ");
        XLSX.writeFile(wb, 'report.xlsx');
      }
    //}
  //});
};

var obj = JSON.parse(fs.readFileSync('A.json', 'utf8'));
converter.json2csv(obj, json2csvCallback);

filename = 'B.csv';
var obj = JSON.parse(fs.readFileSync('B.json', 'utf8'));
converter.json2csv(obj, json2csvCallback);

