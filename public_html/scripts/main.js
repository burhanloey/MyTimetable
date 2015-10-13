/* global XLSX, ko, moment */

var url = "timetable/timetable.xlsx";   // url to file location in server
var workbook;
var filteredTimetable;

/* Subject data model */
function Subject(name) {
    var self = this;
    
    self.name = ko.observable(name);
}

/* Timetable row data model */
function Row() {
    var self = this;
    
    self.columns = ko.observableArray([]);
    
    /* Add column to the row */
    self.addColumn = function(subjectName, rowSpan, isTime) {
        self.columns.push(new Column(subjectName, rowSpan, isTime));
    };
}

/* Timetable column data model */
function Column(subjectName, rowSpan, isTime) {
    var self = this;
    
    self.subjectName = ko.observable(subjectName);
    self.rowSpan = ko.observable(rowSpan);
    self.isTime = ko.observable(isTime);
    
    /* Check if the cell should be highlighted, indicating whether the cell have value */
    self.highlight = ko.computed(function() {
        return self.subjectName().length > 0 && !self.isTime();
    });
}

/* Main view model */
function TimetableViewModel() {
    var self = this;
    
    self.subjects = ko.observableArray([]);
    self.rows = ko.observableArray([]);
    self.day = ko.observable(moment().format("dddd"));
    self.currentClass = ko.observable("No class");
    self.nextClass = ko.observable("No class after this");
    self.timeLeft = ko.observable();
    
    /* Run a timer that check current class for each interval */
    self.tick = function() {
        self.day(moment().format("dddd"));
        self.checkCurrentClass();
    };
    setInterval(self.tick, 1000);   // ms
    
    /* Add row to the table */
    self.addRow = function(row) {
        self.rows.push(row);
    };
    
    /* Add subject to subject list */
    self.addSubject = function() {
        self.subjects.push(new Subject(""));
    };
    
    /* Remove subject from subject list */
    self.removeSubject = function(subject) {
        self.subjects.remove(subject);
        saveSubjects(); // remove from localStorage also
    };
    
    /* Refresh timetable */
    self.refresh = function() {
        $("#error").hide();
        saveSubjects();
        processWorkbook(workbook);
    };
    
    /* Check current class */
    self.checkCurrentClass = function() {
        if (filteredTimetable === undefined) return;
        
        var now = moment();
        var today = now.format("dddd").toUpperCase();
        var currentHour = hourOfDay(now.hour());
        
        if (filteredTimetable.hasOwnProperty(today)) {
            var day = filteredTimetable[today];
            if (day.hasOwnProperty(currentHour)) {
                if (/!merged/.test(day[currentHour].name)) {    // if cell contains word "!merged"
                    self.currentClass(day[currentHour].name.slice("!merged".length));   // remove "!merged"
                } else {
                    self.currentClass(day[currentHour].name);
                }
            } else {
                self.currentClass("No class");
            }
            self.checkNextClass(day, now.hour());
        } else {
            self.currentClass("No class");
        }
    };
    
    /* Check next class */
    self.checkNextClass = function(day, initialHour) {
        for (var hour = initialHour + 1; hour <= 20; hour++) {   // check until 8.00pm
            var nextHour = hourOfDay(hour);
            if (day.hasOwnProperty(nextHour)) {
                if (!/!merged/.test(day[nextHour].name)) {  // if cell does not contains word "!merged"
                    self.nextClass(day[nextHour].name);
                    self.checkTimeLeft(hour);
                    break;
                }
            } else {
                self.nextClass("No class after this");
                self.timeLeft("");
            }
        }
    };
    
    /* Check time left for next class */
    self.checkTimeLeft = function(hour) {
        var now = moment();
        var next = moment().hours(hour).startOf('hour');
        var difference = moment.duration(next.diff(now));
        self.timeLeft("in " + difference.hours() + " hours, " + 
                difference.minutes() + " minutes, " + 
                difference.seconds() + " seconds");
    };
}

/* Apply KnockoutJS bindings */
var timeTable = new TimetableViewModel();
ko.applyBindings(timeTable);

/* Set up XMLHttpRequest */
(function loadWorkbook() {
    var oReq = new XMLHttpRequest();
    oReq.open("GET", url, true);    // timetable url
    oReq.responseType = "arraybuffer";

    oReq.onload = function(e) {
        var arraybuffer = oReq.response;

        /* Convert data to binary string */
        var data = new Uint8Array(arraybuffer);
        var arr = new Array();
        for(var i = 0; i !== data.length; ++i) arr[i] = String.fromCharCode(data[i]);
        var bstr = arr.join("");

        /* Call XLSX */
        workbook = XLSX.read(bstr, {type:"binary"});

        processWorkbook();
        
        $('#loading').fadeOut();    // fade out loading screen after timetable is loaded
    };

    oReq.send();
})();

/* Process workbook and stores data in filteredTimetable */
function processWorkbook() {
    if (workbook === undefined) return;
    
    timeTable.rows.removeAll();
    
    var regex = createRegex();
    if (regex === undefined) return;
    
    filteredTimetable = {MONDAY: {}, TUESDAY: {}, WEDNESDAY: {}, THURSDAY: {}, FRIDAY: {}};
    
    /* Extract information from workbook and store in filteredTimetable */
    for (var attr in filteredTimetable) {
        var worksheet = workbook.Sheets[attr];
        var day = filteredTimetable[attr];
        
        for (var cell in worksheet) {
            if(cell[0] === '!') continue;   // if value is not a cell address
            
            if (regex.test(worksheet[cell].v)) {    // if found subject
                if (worksheet['A' + cell.slice(1)] === undefined) continue; // if class location is undefined
                
                var rowSpanRequired = calcSpan(cell, worksheet);
                var time = worksheet[cell[0] + 1].v;
                var location = worksheet['A' + cell.slice(1)].v;
                
                if (!day.hasOwnProperty(time)) {
                    var subjectName = worksheet[cell].v + " - " + location; // append class location
                    day[time] = {name: subjectName, rowspan: rowSpanRequired};
                    
                    /* Fill next cells to represent merged cells */
                    var currentCell = cell.charCodeAt(0);
                    for (var i = 1; i < rowSpanRequired; i++) {
                        var nextCell = String.fromCharCode(currentCell + i);
                        day[worksheet[nextCell + 1].v] = {name: "!merged" + subjectName, rowspan: 1};
                    }
                } else {
                    $("#error").show(); // show error message if there is a time clash
                }
            }
        }
    }
    
    fillTimetable();
}

/* Create regex from subjects list */
function createRegex() {
    if (timeTable.subjects().length <= 0) return undefined; // if no subjects
    
    /**
     * Tokenize each subjects(for class group), then wrap it with '(?=.*' and ')'.
     * The wrapped strings are then merged with delimiter '|'.
     * 
     * For example: [WIX1002 G4, WIA1001] will become (?=.*WIX1002)(?=.*G4)|(?=.*WIA1001)
     */
    var regexStr = timeTable.subjects()
            .map(function(subject) {
                return subject.name()
                    .split(" ")
                    .reduce(function(subject, token) {
                        return (token.length > 0) ? subject + "(?=.*" + token + ")" : subject;
                    }, "");
            })
            .reduce(function(regex, subject) {
                return (subject.length > 0) ? regex + "|" + subject : regex;
            });
    
    return new RegExp(regexStr);
}

/* Calculate span for cell */
function calcSpan(cell, worksheet) {
    var column = cell.charCodeAt(0) % 65;   // change A to 0, B to 1, and so on...
    var row = parseInt(cell.slice(1)) - 1;
    
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

/* Fill timetable with filteredTimetable */
function fillTimetable() {
    for (var i = 66; i <= 78; i++) {    // from B to N
        var row = new Row();
        var firstWorksheet = workbook.Sheets["MONDAY"];
        var time = firstWorksheet[String.fromCharCode(i) + 1].v;
        row.addColumn(time, 1, true);
        
        for (var attr in filteredTimetable) {
            var day = filteredTimetable[attr];
            if (day.hasOwnProperty(time)) {
                if (!/!merged/.test(day[time].name)) {   // if cell does not contains word "!merged"
                    row.addColumn(day[time].name, day[time].rowspan, false);
                }
            } else {
                row.addColumn("", 1, false);    // add empty cell if no value
            }
        }
        
        timeTable.addRow(row);
    }
}

/* Save subjects list in localStorage */
function saveSubjects() {
    if (timeTable.subjects().length <= 0) { // if no subjects
        localStorage.removeItem("subjects");
        return;
    }
    
    /* Merges subjects name with comma delimiter ',' */
    var subjectList = timeTable.subjects()
            .filter(function(subject) { return subject.name().length > 0; })
            .map(function(subject) { return subject.name(); })
            .reduce(function(list, subject) { return list + "," + subject; });
    
    localStorage.setItem("subjects", subjectList);
}

/* Load subjects list from localStorage */
(function loadSubjects() {
    var subjectList = localStorage.getItem("subjects");
    if (subjectList === null) return;
    
    var temp = [];
    var subjects = subjectList.split(",");
    subjects.forEach(function(subject) {
        temp.push(new Subject(subject));
    });
    timeTable.subjects(temp);
})();

/* Formatting hour according to the original timetable */
function hourOfDay(hourNo) {
    return hourNo + ".00 - " + hourNo + ".59";
}

/* Toggle chevron arrow on hide and show */
function toggleArrow() {
    $('.arrow').toggleClass("glyphicon-chevron-right glyphicon-chevron-down");
}
$('.panel-hide > .panel-body').on('hide.bs.collapse', toggleArrow);
$('.panel-hide > .panel-body').on('show.bs.collapse', toggleArrow);

/* Scroll website to My Subjects section after the list is shown */
$('.panel-hide > .panel-body').on('shown.bs.collapse', function() {
    $('html, body').animate({ scrollTop: $('#my-subjects-panel').offset().top });
});

/* Toggle visibility of panel-hide body */
$('.panel-hide > .panel-heading').click(function() {
    $('.panel-hide > .panel-body').collapse('toggle');
});

/* Set up drag-and-drop event */
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

            /* If binary string, read with type 'binary' */
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

var drop = document.getElementById('drop');
drop.addEventListener('dragenter', handleDragover, false);
drop.addEventListener('dragover', handleDragover, false);
drop.addEventListener('drop', handleDrop, false);
