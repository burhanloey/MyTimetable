function Subject(name) {
    this.name = ko.observable(name);
}

function Row() {
    this.columns = ko.observableArray([]);
    
    this.getLength = ko.computed(function() {
        var lengthOfRow = 0;
        for (var column = 0; column < this.columns().length; column++) {
            lengthOfRow += this.columns()[column].columnSpan();
        }
        return lengthOfRow;
    }, this);
    
    this.setDay = function(day) {
        this.columns()[0].cell(day);
    };
    
    this.addColumn = function(cell, columnSpan) {
        this.columns.push(new Column(cell, columnSpan));
    };
}

function Column(text, columnSpan) {
    this.text = ko.observable(text);
    this.columnSpan = ko.observable(columnSpan);
}

function TimetableViewModel() {
    this.subjects = ko.observableArray([
        new Subject("G4.*WXES1116"),
        new Subject("WMES3302"),
        new Subject("GREK1007")
    ]);
    this.rows = ko.observableArray([]);
    
    this.addRow = function(row) {
        this.rows.push(row);
    };
    
    this.addSubject = function() {
        this.subjects.push(new Subject());
    };
    
    this.refresh = function() {
        processWorkbook(workbook);
    };
}

var timeTable = new TimetableViewModel();
ko.applyBindings(timeTable);

/* set up drag-and-drop event */
var workbook;
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
            workbook = XLSX.read(data, {type: 'binary'});

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
    timeTable.rows.removeAll();
    
    var error = "";
    
    var regex = createRegex();
    
    var sheetNameList = workbook.SheetNames;
    sheetNameList.forEach(function(day) {
        var row = new Row();
        row.addColumn(day, 1);
        
        var worksheet = workbook.Sheets[day];
        for (var cell in worksheet) {
            if (cell[0] === '!') {
                continue;
            }
            if (regex.test(worksheet[cell].v)) {    // if found subjects
                var position = cell.charCodeAt(0) % 65;
                var text = worksheet[cell].v + " - " + worksheet['A' + cell.slice(1)].v;
                
                if (row.getLength() <= position) {
                    for (var i = row.getLength(); i < position; i++) {
                        row.addColumn("", 1);
                    }
                    row.addColumn(text, columnSpan(cell, worksheet));
                } else {
                    error += worksheet[cell].v + " is clashing with some of the classes<br/>";
                    row.columns()[position].text(text);
                    row.columns()[position].columnSpan(columnSpan(cell, worksheet));
                }
            }
        }
        
        /* Fill the rest of row with empty cell */
        for (var column = row.getLength(); column <= 13; column++) {
            row.addColumn("", 1);
        }
        
        timeTable.addRow(row);
    });
    document.getElementById('error').innerHTML = error;
    
    var output = JSON.stringify(to_json(workbook), 2, 2);
    document.getElementById('output').innerHTML = output;
}

function createRegex() {
    var regexStr = "";
    for (var i = 0; i < timeTable.subjects().length; i++) {
        if (i !== 0) {
            regexStr += "|";
        }
        regexStr += timeTable.subjects()[i].name();
    }
    return new RegExp(regexStr);
}

var columnSpan = function(cell, worksheet) {
    var column = cell.charCodeAt(0) % 65;
    var row = parseInt(cell.slice(1));
    
    var ranges = worksheet['!merges'];
    for (var i = 0; i < ranges.length; i++) {
        var startColumn = parseInt(JSON.stringify(ranges[i].s.c));
        var startRow = parseInt(JSON.stringify(ranges[i].s.r));
        var endColumn = parseInt(JSON.stringify(ranges[i].e.c));
        
        if (column === startColumn && row === startRow) {
            return 1 + (endColumn - startColumn);
        }
    }
    return 1;
};

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
