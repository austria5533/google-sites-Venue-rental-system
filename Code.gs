var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();

// Calendars to output appointments to
var cal = CalendarApp.getCalendarById('om7tg43lkklqt43rdhjcrooga4@group.calendar.google.com');

// Create an object from user submission
function Submission(){
  var row = lastRow;
  this.timestamp = sheet.getRange(row, 1).getValue();
  this.name = sheet.getRange(row, 2).getValue();
  this.email = sheet.getRange(row, 3).getValue();
  this.reason = sheet.getRange(row, 4).getValue();
  this.date = sheet.getRange(row, 5).getValue();
  this.time = sheet.getRange(row, 6).getValue();
  this.duration = sheet.getRange(row, 7).getValue();
  this.room = sheet.getRange(row, 8).getValue();
  this.activity = sheet.getRange(row, 9).getValue();
  // Info not from spreadsheet
  this.roomInt = this.room.replace(/\D+/g, '');
  //this.activityInt = this.activity.replace(/\D+/g, '');
  this.status;
  //this.timeString1 =this.time.toTimeString(this.time.getMinutes()+24);
  this.dateString = (this.date.getMonth() + 1) + '/' + this.date.getDate() + '/' + this.date.getYear();
  this.dateString1 = (this.date.getMonth() + 1) + '/' + this.date.getDate();
  this.timeString1 = this.time.getMinutes()+24;
  this.timeString = this.time.toLocaleTimeString(this.timeString1);
  
  if(this.timeString>=36)
  {
    this.date.setHours(this.time.getHours()+1);
  }else
  {
     this.date.setHours(this.time.getHours());
     this.date.setMinutes(this.time.getMinutes()+24);
  }
  
  this.calendar = eval('cal' + String(this.roomInt));
  return this;
}


// Use duration to create endTime variable
function getEndTime(request){
  request.endTime = new Date(request.date);
  switch (request.duration){
    case "30 分鐘":
      request.endTime.setMinutes(request.date.getMinutes() + 30);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "45 分鐘":
      request.endTime.setMinutes(request.date.getMinutes() + 45);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "1 小時":
      request.endTime.setMinutes(request.date.getMinutes() + 60);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "2 小時":
      request.endTime.setMinutes(request.date.getMinutes() + 120);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "3 小時":
      request.endTime.setMinutes(request.date.getMinutes() + 180);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "4 小時":
      request.endTime.setMinutes(request.date.getMinutes() + 240);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "5 小時":
      request.endTime.setMinutes(request.date.getMinutes() + 300);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "6 小時":
      request.endTime.setMinutes(request.date.getMinutes() + 360);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "7 小時":
      request.endTime.setMinutes(request.date.getMinutes() + 420);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "8 小時":
      request.endTime.setMinutes(request.date.getMinutes() + 480);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
     case "9 小時":
      request.endTime.setMinutes(request.date.getMinutes() + 540);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
     case "10 小時":
      request.endTime.setMinutes(request.date.getMinutes() + 600);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
  }
}

// Check for appointment conflicts
function getConflicts(request){
  var conflicts = request.calendar.getEvents(request.date, request.endTime);
  if (conflicts.length < 1) {
    request.status = "Approve";
  } else {
    request.status = "Conflict";
  }
}

function draftEmail(request){
  request.buttonLink = "https://forms.gle/h5e8tb5mhVmN6ruD6";
  request.buttonText = "New Request";
  switch (request.status) {
    case "Approve":
      request.subject = "預約成功: " + request.room + " Reservation for " + request.dateString;
      request.header = "活動空間預約成功";
      request.message = "請好好使用、珍惜活動空間~";
      break;
    case "Conflict":
      request.subject = "預約失敗 " + request.room + "Reservation for " + request.dateString;
      request.header = "活動空間預約失敗";
      request.message = "這個時間有人預約了!請預約其他時段"
      request.buttonText = "重新預約";
      break;
  }
}

function updateCalendar(request){
  var event = request.calendar.createEvent(
    request.name,
    request.date,
    //request.time,
    request.endTime
    )
}

function sendEmail(request){
  MailApp.sendEmail({
    to: request.email,
    subject: request.header,
    htmlBody: makeEmail(request)
  })
  sheet.getRange(lastRow, lastColumn).setValue("Sent: " + request.status);
}

// --------------- main --------------------

function main(){
  var request = new Submission();
  getEndTime(request);
  getConflicts(request);
  draftEmail(request);
  if (request.status == "Approve") updateCalendar(request);
  sendEmail(request);
}