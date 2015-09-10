function Row(row) {
    this.A = row.A === undefined ? "" : row.A;
    this.B = row.B === undefined ? "" : row.B;
    this.C = row.C === undefined ? "" : row.C;
    this.D = row.D === undefined ? "" : row.D;
    this.E = row.E === undefined ? "" : row.E;
    this.F = row.F === undefined ? "" : row.F;
    this.G = row.G === undefined ? "" : row.G;
    this.H = row.H === undefined ? "" : row.H;
    this.I = row.I === undefined ? "" : row.I;
    this.J = row.J === undefined ? "" : row.J;
    this.K = row.K === undefined ? "" : row.K;
    this.L = row.L === undefined ? "" : row.L;
    this.M = row.M === undefined ? "" : row.M;
    this.N = row.N === undefined ? "" : row.N;
}

function TimetableViewModel() {
    this.subjects = ko.observableArray([
        "WXES1116", "WMES3302", "GREK1007"
    ]);
    this.row = ko.observableArray([]);
    
    this.add = function(row) {
        this.row.push(new Row(row));
    };
}

var timeTable = new TimetableViewModel();
ko.applyBindings(timeTable);

/* set up drag-and-drop event */
function handleDrop(e) {
    e.stopPropagation();
    e.preventDefault();
    var files = e.dataTransfer.files;
    var i,f;
    for (i = 0, f = files[i]; i !== files.length; ++i) {
        var reader = new FileReader();
        var name = f.name;
        reader.onload = function(e) {
            var data = e.target.result;

            /* if binary string, read with type 'binary' */
            var workbook = XLSX.read(data, {type: 'binary'});

            processWorkbook(workbook);
        };
        reader.readAsBinaryString(f);
    }
}

function handleDragover(e) {
    e.stopPropagation();
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
}

function processWorkbook(workbook) {
    var regexStr = "";
    for (var i = 0; i < timeTable.subjects().length; i++) {
        if (i !== 0) {
            regexStr += "|";
        }
        regexStr += timeTable.subjects()[i];
    }
    var regex = new RegExp(regexStr, "i");
    
    var firstSheet = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[firstSheet];
    for (var i = 2; i <= 17; i++) {
        var row = {};
        for (var cell in worksheet) {
            if (cell[0] === '!') {
                continue;
            }
            if (cell.charAt(1) === i.toString()) {
                if (regex.test(worksheet[cell].v)) {
                    row['A'] = firstSheet;
                    row[cell.charAt(0)] = worksheet[cell].v;
                    console.log(worksheet[cell].v);
                }
            }
        }
        timeTable.add(row);
    }
    
    var output = JSON.stringify(to_json(workbook), 2, 2);
    document.getElementById('output').innerHTML = output;
}

function to_json(workbook) {
    var result = {};
    workbook.SheetNames.forEach(function(sheetName) {
        var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        if(roa.length > 0){
            result[sheetName] = roa;
        }
    });
    return result;
}

var drop = document.getElementById('drop');
drop.addEventListener('dragenter', handleDragover, false);
drop.addEventListener('dragover', handleDragover, false);
drop.addEventListener('drop', handleDrop, false);
