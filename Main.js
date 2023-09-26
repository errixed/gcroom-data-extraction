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

function assignmentData() {
  var courseId = '524510349937';

  try {
    var classwork = Classroom.Courses.CourseWork.list(courseId);
    var assignments = classwork.courseWork;

    if (assignments.length === 0) {
      Logger.log("No assignments found.");
    } else {
      for (var i = 0; i < assignments.length; i++) {
        var assignment = assignments[i];
        Logger.log("Assignment ID: " + assignment.id);
        Logger.log("Title: " + assignment.title);
        Logger.log("Due Date: " + assignment.dueDate);
        Logger.log("Status: " + assignment.state);
        Logger.log('----------');
      }
    }
  } catch (error) {
    Logger.log('Error: ' + error);
  }
}

function submissionData() {
  var courseId = "524510349937";
  var assignmentId = "524521431584";

  try {
    var students = Classroom.Courses.Students.list(courseId).students;
    var submissions = Classroom.Courses.CourseWork.StudentSubmissions.list(courseId, assignmentId).studentSubmissions;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.clearContents();

    if (submissions && submissions.length > 0) {
      for (var i = 0; i < submissions.length; i++) {
        var submission = submissions[i];

        var student = students.find(function(student) {
          return student.userId === submission.userId;
        });

        var data = [
          student.profile.name.fullName,
          student.profile.emailAddress,
          submission.state,
          submission.creationTime,
          submission.updateTime
        ];
        
        sheet.appendRow(data);
      }
    } else {
      Logger.log('No student submissions found for this assignment.');
    }
  } catch (error) {
    Logger.log('Error: ' + error);
  }
}