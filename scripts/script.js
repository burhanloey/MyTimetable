function DataViewModel(location, morning, evening, night) {
  this.location = location;
  this.morning = morning;
  this.evening = evening;
  this.night = night;
}

function TimetableViewModel() {
  this.data = ko.observableArray([
    new DataViewModel("MM6", "Programming", "DSS", "KO-K")
  ]);
}

ko.applyBindings(new TimetableViewModel());
