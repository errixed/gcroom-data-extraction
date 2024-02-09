function generateMainSheet() {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
  var title = ["classrooms", "selected classrooms", "update time", "send email", "period", "repeat"]
  
  var updateTimeDropdownItems = ["Daily", "Weekly"];
  var updateTimeDropdownCell = mainSheet.getRange('C6');
  var updateTimeDropdownRule = SpreadsheetApp.newDataValidation().requireValueInList(updateTimeDropdownItems).build();
  
  var periodDropdownItems = ["Weekly", "Monthly"];
  var periodDropdownCell = mainSheet.getRange('E6');
  var periodDropdownRule = SpreadsheetApp.newDataValidation().requireValueInList(periodDropdownItems).build();

  var repeatDropdownItems = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
  var repeatDropdownCell = mainSheet.getRange('F6');
  var repeatDropdownRule = SpreadsheetApp.newDataValidation().requireValueInList(repeatDropdownItems).build();  

  var classroomCheckBoxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  var emailCheckBoxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  
  var options = {
    teacherId: 'me',
    courseStates: 'ACTIVE'
  };

  mainSheet.clearContents();
  mainSheet.getRange("A:F").setDataValidation(null);

  var sheetTitle = "GCRoom Extractor";
  var intro = "Rename your first sheet to 'Main'\nGo to 'check classworks' tab.\nChoose 'Generate Main sheet' to generate the setup page. In this page you can choose the classes you want to get update on.\nAfter you setted prefered update time, choose 'Schedule classrooms data' to set a timer to get updates on chosen classes classworks\nAfter you setted prefered sending time, choose 'Schedule send email' to set a timer to send email";
  var footer = "By Amir Shekarabi\nhttps://github.com/errixed\n©2023"
  mainSheet.getRange("A1").setValue(sheetTitle).setFontWeight("bold").setFontSize("25");
  mainSheet.getRange("A1:F1").setBackground("lightblue");
  mainSheet.getRange("A2").setValue(intro).setFontSize("12");
  mainSheet.getRange("A2:F2").setBackground("yellow");
  mainSheet.getRange("A3").setValue(footer).setFontWeight("bold").setFontSize("10");
  mainSheet.getRange("A3:F3").setBackground("lightgreen");

  mainSheet.getRange(5, 1, 1, 6).setValues([title]).setFontWeight("bold");

  updateTimeDropdownCell.setDataValidation(updateTimeDropdownRule).setFontColor("black");
  periodDropdownCell.setDataValidation(periodDropdownRule).setFontColor("black");
  repeatDropdownCell.setDataValidation(repeatDropdownRule).setFontColor("black");
  
  var lastRow = mainSheet.getLastRow();
  var lastColumn = mainSheet.getLastColumn();
  for (var row = 1; row <= lastRow; row++) {
    for (var column = 1; column <= lastColumn; column++) {
      mainSheet.setColumnWidth(column, 215);
    }
  }

  try {
    var courseList = Classroom.Courses.list(options);
    var classroomCheckBoxTargetCell;
    var classroomCheckboxCell;
    var emailCheckBoxTargetCell;
    var emailCheckBoxCell;
    
    if (courseList && Array.isArray(courseList.courses)) {
      courseList.courses.forEach((courseItem, index) => {
        mainSheet.appendRow([courseItem.name]);
        
        classroomCheckBoxTargetCell = mainSheet.getRange(`A${index+6}`);
        classroomCheckboxCell = classroomCheckBoxTargetCell.offset(0, 1);
        classroomCheckboxCell.setDataValidation(classroomCheckBoxRule).setFontColor("grey");

        emailCheckBoxTargetCell = mainSheet.getRange(`D${index+6}`);
        emailCheckBoxCell = emailCheckBoxTargetCell.offset(0, 0);
        emailCheckBoxCell.setDataValidation(emailCheckBoxRule).setFontColor("grey");
      });
    }
  } catch (error) {
    throw new Error(error);
  }

  updateTimeDropdownCell.setValue("Daily");
  periodDropdownCell.setValue("Weekly");
  repeatDropdownCell.setValue("Tuesday");
}

function triggerClassroomUpdate() {
  deleteTriggerByFunctionName("classroomUpdate")
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
  var updateTimeDropdownCell = mainSheet.getRange('C6').getValue();

  if (updateTimeDropdownCell == "Daily") {
    ScriptApp.newTrigger("classroomUpdate")
      .timeBased()
      .everyDays(1)
      .atHour(6)
      .create();
  } else if (updateTimeDropdownCell == "Weekly") {
    ScriptApp.newTrigger("classroomUpdate")
      .timeBased()
      .everyWeeks(1)
      .create();
  }
}

function classroomUpdate() {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  var classrooms = mainSheet.getRange("A6:A").getValues();
  var classroomsFilter = classrooms.filter(function (x) {
    return !(x.every(element => element === ""))
  });
  var classroomCheckbox = mainSheet.getRange('B6:B').getValues();
  var classroomCheckboxFilter = classroomCheckbox.filter(function (x) {
    return !(x.every(element => element === ""))
  });
  var selectedClassroom = [];
  for (var i = 0; i < classroomCheckboxFilter.length; i++) {
    if (classroomCheckboxFilter[i] == "true") {
      selectedClassroom.push(classroomsFilter[i]);
    }
  }

  sheets.forEach(sheet => {
    if (sheet.getName() !== "Main") {
      spreadsheet.deleteSheet(sheet);
    }
  });

  selectedClassroom.forEach(classroom => {
    spreadsheet.insertSheet(classroom.toString()); 
    generateClassroomData(findCourseIdByName(classroom.toString()));
  });

}

function generateClassroomData(id) {
  var courseId = [id];
  Logger.log(courseId)

  var assignments = Classroom.Courses.CourseWork.list(courseId).courseWork;
  var students = Classroom.Courses.Students.list(courseId).students;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  sheet.clearContents();

  var title = ["name", "email"];
  var studentName = [];
  var studentEmail = [];
  var submissionState = [];
  var assignmentTitle = [];

  for (var i = 0; i < assignments.length; i++) {
    var assignment = assignments[i];
    var submissions = Classroom.Courses.CourseWork.StudentSubmissions.list(courseId, assignment.id).studentSubmissions;

    for (var j = 0; j < submissions.length; j++) {
      var submission = submissions[j];
      var student = students.find(function(student) {
        return student.userId === submission.userId;
      });
      
      studentName.push(student.profile.name.fullName);
      studentEmail.push(student.profile.emailAddress);
      submissionState.push(submission.state);
    }
    assignmentTitle.push(assignment.title)
    title.push(assignment.title);
  }

  sheet.appendRow(title)
  var lastRow = sheet.getLastRow() + 1;
  var eachStudentName = studentName.filter((item, index) => studentName.indexOf(item) === index);
  
  for (var i = 0; i < eachStudentName.length; i++) {
    sheet.getRange(lastRow + i, 1).setValue(studentName[i]);
  }

  for (var i = 0; i < eachStudentName.length; i++) {
    sheet.getRange(lastRow + i, 2).setValue(studentEmail[i]);
  }

  var finalSubmissionState = [];
  for (var i = 0; i < submissionState.length; i += eachStudentName.length) {
    finalSubmissionState.push(submissionState.slice(i, i + eachStudentName.length));
  }

  var stateLastRows = [];
  for (var i = 1; i <= finalSubmissionState.length; i++) {
    stateLastRows.push(lastRow);
  }

  for (var col = 0; col < finalSubmissionState[0].length; col++) {
    for (var row = 0; row < finalSubmissionState.length; row++) {
      var insertedData = sheet.getRange(stateLastRows[row] + col, row + 3).setValue(finalSubmissionState[row][col]);
      if (finalSubmissionState[row][col] == "TURNED_IN") {
        insertedData.setFontColor("green");
      } else if (finalSubmissionState[row][col] == "NEW") {
        insertedData.setFontColor("blue");
      } else {
        insertedData.setFontColor("red");
      }
    }
  }

  sheet.getRange(1, 1, 1, title.length).setValues([title]).setFontWeight("bold");

  Logger.log("EXECUTED SUCCESSFULLY")
}

function triggerGetDataReport() {
  deleteTriggerByFunctionName("getDataReport")
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
  var periodDropdownCell = mainSheet.getRange('E6').getValue();
  var repeatDropdownCell = mainSheet.getRange('F6').getValue();

  if (periodDropdownCell == "Weekly") {
    ScriptApp.newTrigger("getDataReport")
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(weekday(repeatDropdownCell.toString().toLowerCase()))
      .create()
  } else if (periodDropdownCell == "Monthly") {
    ScriptApp.newTrigger("getDataReport")
      .timeBased()
      .onMonthDay(1)
      .create()
  }
}

function getDataReport() {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");

  var classrooms = mainSheet.getRange("A6:A").getValues();
  var classroomsFilter = classrooms.filter(function (x) {
    return !(x.every(element => element === ""))
  });
  var emailCheckbox = mainSheet.getRange('D6:D').getValues();
  var emailCheckboxFilter = emailCheckbox.filter(function (x) {
    return !(x.every(element => element === ""))
  });
  var selectedClassroom = [];
  for (var i = 0; i < emailCheckboxFilter.length; i++) {
    if (emailCheckboxFilter[i] == "true") {
      selectedClassroom.push(classroomsFilter[i]);
    }
  }

  for (var i = 0; i < selectedClassroom.length; i++) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(selectedClassroom[i]); 
    if (sheet.getName() != "Main") {
      // Logger.log(sheet.getName())
      dataReport(sheet)
    }
  }
}

function dataReport(sheet) {

  var data = sheet.getDataRange().getValues();
  var numRows = data.length;
  var numCols = data[0].length;
  var columnData = [];
  for (var col = 0; col < numCols; col++) {
    for (var row = 0; row < numRows; row++) {
      var cellValue = data[row][col];
      columnData.push(cellValue);
    }
  }
  var eachColumnData = [];
  for (var i = 0; i < columnData.length; i += numRows) {
    eachColumnData.push(columnData.slice(i, i + numRows));
  }
  
  var messageData = [];
  var names = [];
  for (var i = 0; i < eachColumnData.length; i ++) {
    for (var j = 0; j < eachColumnData[i].length; j++) {
      names.push(eachColumnData[0][j])
      if (["CREATED", "NEW", "RECLAIMED_BY_STUDENT"].includes(eachColumnData[i][j])) {
        messageData.push(eachColumnData[0][j], eachColumnData[1][j], eachColumnData[i][0])
      }
    }
  }

  var eachMessageData = [];
  for (var i = 0; i < messageData.length; i += 3) {
    eachMessageData.push(messageData.slice(i, i + 3));
  }

  var groupedMessages = [];
  eachMessageData = eachMessageData.forEach(arr => {
    var key = arr[0];
    if (!groupedMessages[key]) {
      groupedMessages[key] = [];
    }
    groupedMessages[key].push(arr);
  })
  var flatedGroupedMessages = Object.values(groupedMessages).flat();

  var mergedArrays = [];
  for (var subArray of flatedGroupedMessages) {
    var key = subArray[0];
    if (mergedArrays[key]) {
      mergedArrays[key] = [...new Set(mergedArrays[key].concat(subArray))];
    } else {
      mergedArrays[key] = subArray;
    }
  }
  var message = Object.values(mergedArrays);

  var emailWords = "This email is a reminder that you have some outstanding work on the Google Classroom to finish:";
  var stEmail;
  var stName;
  var stWork;
  var emailBody;
  for (var i = 0; i < message.length; i++) {
    if (message[i].length >= 2) {
      stName=message[i][0];
      emailBody = "Hello "+stName+", \n\n"+emailWords;
      stEmail=message[i][1];
      stWork =message[i].slice(2);
      emailBody+= stWork;
    }

    Logger.log(emailBody);
    Logger.log(stEmail);

    MailApp.sendEmail(stEmail, "Outstanding Work", emailBody);
  }

}

function weekday(day) {
  if (day == "sunday") {
    return ScriptApp.WeekDay.SUNDAY
  } else if (day == "monday") {
    return ScriptApp.WeekDay.MONDAY
  } else if (day == "tuesday") {
    return ScriptApp.WeekDay.TUESDAY
  } else if (day == "wednesday") {
    return ScriptApp.WeekDay.WEDNESDAY
  } else if (day == "thursday") {
    return ScriptApp.WeekDay.THURSDAY
  } else if (day == "friday") {
    return ScriptApp.WeekDay.FRIDAY
  } else if (day == "saturday") {
    return ScriptApp.WeekDay.SATURDAY
  } else {
    return null
  }
}

function findCourseIdByName(courseName) {
  var courseId = null;
  var courses = Classroom.Courses.list().courses;
  for (var i = 0; i < courses.length; i++) {
    if (courses[i].name == courseName) {
      courseId = courses[i].id;
      break;
    }
  }
  return courseId;
}

function deleteTriggerByFunctionName(functionNameToDelete) {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    var trigger = allTriggers[i];
    if (trigger.getHandlerFunction() === functionNameToDelete) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Check Classworks")
      .addItem("Generate Main sheet", "generateMainSheet")
      .addItem("Generate classrooms data", "classroomUpdate")
      .addItem("Schedule data extraction", "triggerClassroomUpdate")
      .addItem("Schedule sending emails", "triggerGetDataReport")
      .addToUi();
}