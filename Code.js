// IMPORTANT; MIT LICENSE attached. Copyright (c) 2023 Amir H. Shekarabi. Free to use by mentioning the source
// IMPORTANT; create a google sheet, go to extension tab, select apps script. run the code from there
// IMPORTANT; appsscript.json file is required with all the codes in it

// courseData function will return all courses IDs
function courseData() {
  const arguments = {
    teacherId: 'me',
    courseStates: 'ACTIVE'
  };

  try {
    const course = Classroom.Courses.list(arguments).courses
    for(let i = 0; i < course.length; i++){
      Logger.log("course name: " + course[i].name)
      Logger.log("course ID: " + course[i].id)
    }
  } catch (error) {
    Logger.log('Error: ' + error);
  }
}

// assignmentSubmissionState function will insert all student names, student emails and status of submission of every assignment in the selected course, to the created google sheet
function insertData() {

  // replace SELECTED_COURSE_ID with your course ID of choice
  var courseId = '524510349937';

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
  
  // below lines will insert student names to first column of the created google sheet
  for (var i = 0; i < eachStudentName.length; i++) {
    sheet.getRange(lastRow + i, 1).setValue(studentName[i]);
  }
  // below lines will insert student emails to second column of the created google sheet
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

  // below lines will insert sumbmission status of all the assigments in the selected course to the created google sheet
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

  // below line will insert all assignments title in the selected selected course to the created google sheet
  sheet.getRange(1, 1, 1, title.length).setValues([title]).setFontWeight("bold");

  // if code execute successfully, you will see this message in the terminal
  Logger.log("EXECUTED SUCCESSFULLY")
}

// notifier function will check students one by one and tell which one didn't do their assignments 
function checkData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

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
      if (["CREATED", "NEW"].includes(eachColumnData[i][j])) {
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
  
  for (var i = 0; i < message.length; i++) {
    Logger.log(message[i]);
  }
}