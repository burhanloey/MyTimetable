function DataViewModel(location, morning, evening, night) {
    this.location = location;
    this.morning = morning;
    this.evening = evening;
    this.night = night;
}

function TimetableViewModel() {
    this.data = ko.observableArray([
        new DataViewModel("MM6", "Programming", "DSS", "KO-K"),
        new DataViewModel("BT4", "DSS", "Programming", "KO-K")
    ]);
}

ko.applyBindings(new TimetableViewModel());

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
    var cellAddress = 'B1';
    var firstSheet = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[firstSheet];
    var cell = worksheet[cellAddress];
    document.getElementById('preview').innerHTML = cellAddress + ': ' + cell.v;
    
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
