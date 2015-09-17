/* global XLSX, ko */

function Subject(name) {
    var self = this;
    
    self.name = ko.observable(name);
}

function Row() {
    var self = this;
    
    self.columns = ko.observableArray([]);
    
    self.addColumn = function(subjectName, rowSpan, isTime) {
        self.columns.push(new Column(subjectName, rowSpan, isTime));
    };
}

function Column(subjectName, rowSpan, isTime) {
    var self = this;
    
    self.subjectName = ko.observable(subjectName);
    self.rowSpan = ko.observable(rowSpan);
    self.isTime = ko.observable(isTime);
    self.highlight = ko.computed(function() {
        return self.subjectName().length > 0 && !self.isTime();
    });
}

function TimetableViewModel() {
    var self = this;
    
    self.subjects = ko.observableArray([
        new Subject("G4 WXES1116"),
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

            processWorkbook();
        };
        reader.readAsBinaryString(f);
    }
}

function handleDragover(e) {
    e.stopPropagation();
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
}

var filteredTimetable;
function processWorkbook() {
    timeTable.rows.removeAll();
    
    var regex = createRegex();
    
    filteredTimetable = {MONDAY: {}, TUESDAY: {}, WEDNESDAY: {}, THURSDAY: {}, FRIDAY: {}};
    
    /* extract information from workbook */
    for (var attr in filteredTimetable) {
        var worksheet = workbook.Sheets[attr];
        var day = filteredTimetable[attr];
        
        for (var cell in worksheet) {
            if(cell[0] === '!') continue;
            
            if (regex.test(worksheet[cell].v)) {    // if found subject
                var rowSpanRequired = calcRowSpan(cell, worksheet);
                var time = worksheet[cell[0] + 1].v;
                var location = worksheet['A' + cell.slice(1)].v;
                
                if (!day.hasOwnProperty(time)) {
                    var subjectName = worksheet[cell].v + " - " + location;
                    day[time] = {name: subjectName, rowspan: rowSpanRequired};
                    
                    /* fill next cells to represent merged cells */
                    var currentCell = cell.charCodeAt(0);
                    for (var i = 1; i < rowSpanRequired; i++) {
                        var nextCell = String.fromCharCode(currentCell + i);
                        day[worksheet[nextCell + 1].v] = {name: "merged", rowspan: 1};
                    }
                } else {
                    $("#error").show();
                }
            }
        }
    }
    
    fillTimetable();
}

function createRegex() {
    var regexStr = "";
    timeTable.subjects().forEach(function(subject, index) {
        if (index !== 0) regexStr += "|";
        
        var tokens = subject.name().split(" ");
        tokens.forEach(function(token) {
            regexStr += "(?=.*" + token + ")";
        });
    });
    return new RegExp(regexStr);
}

function calcRowSpan(cell, worksheet) {
    var column = cell.charCodeAt(0) % 65;
    var row = parseInt(cell.slice(1));
    
    var ranges = worksheet['!merges'];
    for (var index in ranges) {
        var startColumn = parseInt(JSON.stringify(ranges[index].s.c));
        var startRow = parseInt(JSON.stringify(ranges[index].s.r));
        var endColumn = parseInt(JSON.stringify(ranges[index].e.c));
        
        if (column === startColumn && row === startRow) {
            return 1 + (endColumn - startColumn);
        }
    }
    return 1;
};

function fillTimetable() {
    for (var i = 66; i <= 78; i++) {    // from B to N
        var row = new Row();
        var firstWorksheet = workbook.Sheets["MONDAY"];
        var time = firstWorksheet[String.fromCharCode(i) + 1].v;
        row.addColumn(time, 1, true);
        
        for (var attr in filteredTimetable) {
            var day = filteredTimetable[attr];
            if (day.hasOwnProperty(time)) {
                if (day[time].name !== "merged") {
                    row.addColumn(day[time].name, day[time].rowspan, false);
                }
            } else {
                row.addColumn("", 1, false);
            }
        }
        
        timeTable.addRow(row);
    }
}

var drop = document.getElementById('drop');
drop.addEventListener('dragenter', handleDragover, false);
drop.addEventListener('dragover', handleDragover, false);
drop.addEventListener('drop', handleDrop, false);
