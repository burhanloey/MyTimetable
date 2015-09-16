/* global XLSX, ko */

function Subject(name) {
    var self = this;
    
    self.name = ko.observable(name);
}

function Row() {
    var self = this;
    
    self.columns = ko.observableArray([]);
    
    self.getLength = ko.computed(function() {
        var lengthOfRow = 0;
        for (var column = 0; column < self.columns().length; column++) {
            lengthOfRow += self.columns()[column].columnSpan();
        }
        return lengthOfRow;
    });
    
    self.setDay = function(day) {
        self.columns()[0].cell(day);
    };
    
    self.addColumn = function(cell, columnSpan) {
        self.columns.push(new Column(cell, columnSpan));
    };
}

function Column(text, columnSpan) {
    var self = this;
    
    self.text = ko.observable(text);
    self.columnSpan = ko.observable(columnSpan);
}

function TimetableViewModel() {
    var self = this;
    
    self.subjects = ko.observableArray([
        new Subject("G4.*WXES1116"),
//        new Subject("WXES1116"),
        new Subject("WMES3302"),
        new Subject("GREK1007")
    ]);
    self.rows = ko.observableArray([]);
    
    self.addRow = function(row) {
        self.rows.push(row);
    };
    
    self.addSubject = function() {
        self.subjects.push(new Subject());
    };
    
    self.removeSubject = function(subject) {
        self.subjects.remove(subject);
    };
    
    self.refresh = function() {
        $("#error").hide();
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
    
    var regex = createRegex();
    
    var filteredTimetable = {MONDAY: {}, TUESDAY: {}, WEDNESDAY: {}, THURSDAY: {}, FRIDAY: {}};
    
    var sheetNameList = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"];
    sheetNameList.forEach(function(day) {
        var worksheet = workbook.Sheets[day];
        for (var cell in worksheet) {
            if(cell[0] === '!') continue;
            if (regex.test(worksheet[cell].v)) {
//                console.log(day + "!" + cell + "=" + JSON.stringify(worksheet[cell].v));
                filteredTimetable[day][cell[0] + 1] = {name: worksheet[cell].v, rowspan: calcRowSpan(cell, worksheet)};
            }
        }
    });
    
    for (var day in filteredTimetable) {
        var dayObj = filteredTimetable[day];
        for (var time in dayObj) {
            var timeObj = dayObj[time];
            console.log("subject: " + timeObj.name + ", span: " + timeObj.rowspan);
        }
    }
    
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

var calcRowSpan = function(cell, worksheet) {
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
