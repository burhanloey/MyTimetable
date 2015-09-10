function UserSubjects() {
    this.subjects = ko.observableArray([
        "WMES3302", "WXES1116", "GREK1007"
    ]);
}

function RowViewModel(row) {
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

var TimetableViewModel = function() {
    this.data = ko.observableArray([]);
    
    this.addData = function(row) {
        this.data.push(new RowViewModel(row));
    }
}

var timeTable = new TimetableViewModel();
ko.applyBindings(timeTable);

/* set up drag-and-drop event */
function handleDrop(e) {
    e.stopPropagation();
    e.preventDefault();
    var files = e.dataTransfer.files;
    var i,f;
    for (i = 0, f = files[i]; i != files.length; ++i) {
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
    var firstSheet = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[firstSheet];
    for (var i = 2; i <= 17; i++) {
        var row = {};
        for (z in worksheet) {
            if (z[0] === '!') {
                continue;
            }
            if (z.charAt(1) == i.toString()) {
                var regex = /WXES2114/;
                var regexNoRoom = /W|^G/;
                if (regex.test(worksheet[z].v) || !regexNoRoom.test(worksheet[z].v)) {
                    row[z.charAt(0)] = worksheet[z].v;
                }
            }
        }
        timeTable.addData(row);
    }
    document.getElementById('preview').innerHTML = JSON.stringify(worksheet['!merges'], 2, 2);
    
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
