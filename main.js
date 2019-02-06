/*

OVERVIEW:

This script will read in data from a spreadsheet and generate custom spreadsheets
according to specific needs. To use this script with a different set of start/end
dates, different output email target, or a different spreadsheet source, you will
need to update the values of the variables declared in the 'Global vars' section 
below.

Created by Kevin Birk in Jan '19

*/

function generateSheets() {
  // Top of program
  //console.log('starting main');
  
  // Global vars
  //console.log('global vars');
  var startDate = 'DD/MM/YY'; // ie '01/01/18'
  var endDate = 'DD/MM/YY';
  var emailOutputTarget = 'example@domain.com';
  // Get from the URL of the target sheet
  var sheetId =  'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
  /* These values below are received from an external program.
  When data is read in from each row, the ID will be checked against
  this array. If it is found, then part of the ID number will be sliced off,
  merging the sections together in a single ID-entry. */
  var comboCourses = ['PROGR-UG.0001.001',
                      'PROGR-UG.0001.002',
                      'PROGR-UG.0001.001',
                      'PROGR-UG.0001.002',];
  // End global vars
    
  // Data Objects
  // console.log('data objects');
  var users = [];
  var courses = [];
  var coursesUsers = [];
  
  // Reading and preparing the data
  // console.log('reading and preparing data');
  var targetSheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  
  // Measure range of sheet, find header row start and data row start
  // console.log('measuring range');
  var range = targetSheet.getDataRange();
  var table = range.getValues();
  var headerStart = findHeaderRowStart(range);
  var rowStart = headerStart + 2;
  
  // Map column names and set column numbers of needed data properties
  // userId, fullName; courseId, courseName; courseId, userId
  // console.log('mapping columns');
  var colNames = mapColumns(range, headerStart);
  var colUserId = colNames.indexOf('UserID');
  var colFullName = colNames.indexOf('Full Name');
  var colCourseId = colNames.indexOf('Course ID');
  var colCourseName = colNames.indexOf('Course Title');
  var colSectionType = colNames.indexOf('Section Type');
  
  // Sort sheet alphabetically by userId
  // console.log('sort sheet alpha');
  var range = targetSheet.getRange(rowStart, 1, (range.getHeight() - rowStart), range.getWidth());
  range.sort(colUserId + 1); // +1 when using range rather than rangeValues (js starts at 0; apps script at 1)
  
  // Create New Sheets to place data into
  // console.log('create new sheets');
  
  // Define vars
  var tempNewParent = DriveApp.getFileById(sheetId).getParents().next().getId(); // location of source data
  
  // Create sheet - Users
  var usersSheet = SpreadsheetApp.create('examtakers', 2, 8);
  var usersSheetId = usersSheet.getId(); // New
  // Move created sheet into location of source data - Users
  // var tempFileId = usersSheet.getId(); // Old
  var tempOldParent = DriveApp.getFileById(usersSheetId).getParents().next().getId();
  moveFile(usersSheetId, tempNewParent, tempOldParent);
  
  // Format sheet - Users
  usersSheet = usersSheet.getSheets()[0];
  var usersColHeaders = [
    ['User ID', 'Lname', 'Fname', 'Email', 'Passwd', 'LabEquip', 'Group #', 'External Id']
  ];
  
  // Create sheet - Courses
  var coursesSheet = SpreadsheetApp.create('courses', 2, 7);
  var coursesSheetId = coursesSheet.getId(); // New
  // Move created sheet into location of source data
  // var tempFileId = coursesSheet.getId(); - Old
  var tempOldParent = DriveApp.getFileById(coursesSheetId).getParents().next().getId();
  moveFile(coursesSheetId, tempNewParent, tempOldParent);
  
  // Format sheet - Courses
  coursesSheet = coursesSheet.getSheets()[0];
  var coursesColHeaders = [
    ['Course ID', 'Course Name', 'Instructor Lname', 'Instructor Fname', 'Instructor Title', 'StartDate', 'EndDate']
  ];
  
  // Create sheet - CoursesUsers
  var coursesUsersSheet = SpreadsheetApp.create('coursesexamtakers', 2, 2);
  var coursesUsersSheetId = coursesUsersSheet.getId();
  // Move created sheet into location of source data
  // var tempFileId = coursesUsersSheet.getId();
  var tempOldParent = DriveApp.getFileById(coursesUsersSheetId).getParents().next().getId();
  moveFile(coursesUsersSheetId, tempNewParent, tempOldParent);
  
  // Format sheet - CoursesUsers
  coursesUsersSheet = coursesUsersSheet.getSheets()[0];
  var coursesUsersColHeaders = [
    ['CourseID', 'ExamTakerID']
  ];
  
  // Vars for sheet manipulation
  console.log('vars for sheet manipulation');
  // Users
  var usersRange = usersSheet.getRange(1,1,1,8);
  usersRange.setValues(usersColHeaders);
  var usersRangeStart = 2;
  // Courses
  var coursesRange = coursesSheet.getRange(1,1,1,7);
  coursesRange.setValues(coursesColHeaders);
  var coursesRangeStart = 2;
  // CoursesUsers
  var coursesUsersRange = coursesUsersSheet.getRange(1,1,1,2);
  coursesUsersRange.setValues(coursesUsersColHeaders);
  var coursesUsersRangeStart = 2;
  
  // Read each row of sheet and populate data sets
  console.log('read each row and populate data output')
  var rangeValues = range.getValues();
  for (var count = 0; count < (range.getHeight() - 1); count++) {
    
    // Uncomment to skip data processing when testing other parts of the script
    // break;
    
    // Read row data into temp variables for use in the script
    var tempUserId = rangeValues[count][colUserId];
    var tempFullName = rangeValues[count][colFullName];
    var tempCourseId = formatCourseId(rangeValues[count][colCourseId]);
    var tempCourseName = rangeValues[count][colCourseName];
    
    /* Section Type determines whether or not row is relevant
    Only lecture type courses have data imported*/
    var tempSectionType = rangeValues[count][colSectionType];
    var isTypeLecture;
    if (tempSectionType === 'Lecture') {
      isTypeLecture = 1;
    } else {
      isTypeLecture = 0;
    }
    
    /* Clean tempCourseId- use str.replace with regex matching double periods
    to ensure that bad data was cleaned from system */
    if (isTypeLecture === 0) {
      ; // Nothing happens
    } else {
      tempCourseId = tempCourseId.replace(/\.\./, '.');
    }    
    
    /* Check to see if course value is a combo course.
    If it is a combo course, then change the value to the respective combo course value.
    If not, leave the temp value unchanged.*/
    if (isTypeLecture === 0) {
      ; // Nothing happens
    } else {
      if (searchArray(comboCourses, tempCourseId) === 1) {
        console.log(tempFullName);
        console.log('before: ' + tempCourseId) ;
        tempCourseId = tempCourseId.slice(0, (tempCourseId.length - 4));
        console.log('after: ' + tempCourseId);
      }
    }
    
    // New Method
    // Enclose the below block in an if statement testing if row is of type 'Lecture'
    // If not, the whole row can be skipped
    if (isTypeLecture === 0) {
      ; // Nothing happens
    } else {
      /* Create user, course and courseUser rows:
      In this section, the temp data read from the row is added to each of the new sheets.
      Because the user and courses sheets use unique values, each new entry is checked for
      uniqueness against an array of all user or course ids recorded up to that point. If 
      the target is found in the array of recorded vals, then the user or course is skipped.
      If it is not found, it is written to the respective list, as well as to the array of
      unique values. */

      // Users
      if (searchArray(users, tempUserId) === 1) {
        ; // Nothing happens
      } else {
        // Add user to users array
        users.push(tempUserId);
        // Add row of data to users sheet
        var tempUser = [
          [tempUserId, getLastName(tempFullName), getFirstName(tempFullName), (tempUserId + '@example.com'), Math.random().toString(36).substr(2,8), 0, 0, (tempUserId + '@example.com')]
          ];
        usersRange = usersSheet.getRange(usersRangeStart,1,1,8);
        usersRange.setValues(tempUser);
        usersRangeStart++;
      }
      // Courses
      if ((searchArray(courses, tempCourseId)) === 1) {
        ; // Nothing happens
      } else {
        // Add course to courses array
        courses.push(tempCourseId);
        // Add row to courses sheet
        // Normal course
        var tempCourse = [
          [tempCourseId, tempCourseName, 'Last', 'First', 'Title', startDate, endDate]
          ];
        coursesRange = coursesSheet.getRange(coursesRangeStart,1,1,7);
        coursesRange.setValues(tempCourse);
        coursesRangeStart++;
        // Special course
        // tempCourse[0] = tempCourseId + ' Special';
        var tempCourse = [
          [tempCourseId + ' Special', tempCourseName + ' Special', 'Last', 'First', 'Title', startDate, endDate]
          ];
        coursesRange = coursesSheet.getRange(coursesRangeStart,1,1,7);
        coursesRange.setValues(tempCourse);
        coursesRangeStart++;
        // MakeUp course
        // tempCourse[0] = tempCourseId + ' MakeUp';
        var tempCourse = [
          [tempCourseId + ' MakeUp', tempCourseName + ' MakeUp', 'Last', 'First', 'Title', startDate, endDate]
          ];
        coursesRange = coursesSheet.getRange(coursesRangeStart,1,1,7);
        coursesRange.setValues(tempCourse);
        coursesRangeStart++;
      }
    
      // CoursesUsers
      var tempCourseUser = [
        [tempCourseId, tempUserId]
        ];
      coursesUsersRange = coursesUsersSheet.getRange(coursesUsersRangeStart,1,1,2);
      coursesUsersRange.setValues(tempCourseUser);
      coursesUsersRangeStart++;
    }
  }
  
  // Send email of finished sheets section
  /*
  Need to use advanced services here, which will allow use of the web API to fetch complex files/MIMEtypes
  https://developers.google.com/apps-script/guides/services/advanced#enabling_advanced_services
  */
  
  // Get ID's of created sheets
  
  // Old - does not work - conversion to any mime type besides PDF fails...
  // var usersSheetFile = DriveApp.getFileById(usersSheetId).getAs(MimeType.GOOGLE_SHEETS);
  // var coursesSheetFile = DriveApp.getFileById(coursesSheetId).getAs(MimeType.GOOGLE_SHEETS);
  // var coursesUsersSheetFile = DriveApp.getFileById(coursesUsersSheetId).getAs(MimeType.GOOGLE_SHEETS);
  
  // New - these work but when attached, will only provide the PDF
  // var usersSheetFile = DriveApp.getFileById(usersSheetId);
  // var coursesSheetFile = DriveApp.getFileById(coursesSheetId);
  // var coursesUsersSheetFile = DriveApp.getFileById(coursesUsersSheetId);
  
  // Build vars to pass to func
  
  // Old code for send as attachments
  // var subjectTxt = 'Secure Testing App Prep Data Output Complete';
  // var bodyTxt = 'Attached below is the output data from the Secure Testing App Sheet Preparation script\n'
    + '\n' + 'Regards,\n' + 'Secure Testing App Admin';
  // var  tempAttchmnts = [usersSheetFile, coursesSheetFile, coursesUsersSheetFile];
  // var tempObj = new Object();
  // tempObj.attachments = tempAttchmnts;
  // tempObj.name = 'Secure Testing App Prep Data Output';
  
  // New code for send link
  var folderUrl = DriveApp.getFolderById(tempNewParent).getUrl();
  var subjectTxt = 'Secure Testing App Prep Data Output Complete';
  var bodyTxt = 'The Secure Testing App data prep script has finished.\n'
    + 'Here is a link to the folder where the output files can be found:\n'
    + folderUrl + '\n'
    + 'Regards,\n'
    + 'Secure Testing App Admin';
  
  // Send email
  GmailApp.sendEmail(emailOutputTarget, subjectTxt, bodyTxt);
  
  // End the program
  console.log('end program')
  return;
  
  // Utility functions
  
  // Find Header Row Start - gets the start row for the sheet
  function findHeaderRowStart(range) {
    var count = 0;
    var temp;
    while(true) {
      temp = range.getValues()[count][0];
      console.log(count + ' ' + temp);
      if (temp === 'Term') {
        break;
      }
      count++;
    }
    return count;
  }

  // Map Columns
  function mapColumns(range, headerStart) {
    console.log('mapColumns started');
    var columnNames = [];
    var data = range.getValues();
    for (count = 0; count < (range.getWidth()); count++) {
      columnNames.push(data[headerStart][count]);
    }
    // For testing - use to check that header cols were mapped correctly
    // for each (var colName in columnNames) {
    //   console.log(columnNames.indexOf(colName) + ': ' + colName);
    //   }
    return columnNames;
  }
  
  // Get First Name
  function getFirstName(fullName) {
    var nameArr = fullName.split(/, /);
    var firstName = (nameArr[1]).split(/ /)[0];
    return firstName;
  }
  
  // Get Last Name
  function getLastName(fullName) {
    var re3 = fullName.search(/, /); // find index of char after last char of last name
    return (fullName.slice(0, re3));
  }
  
  // Search Array
  function searchArray(array, target) {
    var wasFound = 0;
    for each (var item in array) {
      if (item === target) {
        wasFound = 1;
        break;
      }
    }
    return wasFound;
  }
  
  // Format Course Id
  function formatCourseId(courseId) {
    var txt = courseId;
    // Old method - no leading zero fill on course codes w/length < 4
    // var newTxt = txt.slice(0,8) + '.' + txt.slice(8, (txt.length - 3)) + '.' + txt.slice((txt.length - 3));
    // New method that will fill with leading zeroes for the course code portion
    var tmpTxt1 = txt.slice(0,8);
    var tmpTxt2 = txt.slice(8,(txt.length - 3));
    var cntLimit = (4 - tmpTxt2.length);
    for (var cnt = 0; cnt < cntLimit; cnt++){
      tmpTxt2 = '0' + tmpTxt2;
    }
    var tmpTxt3 = txt.slice((txt.length - 3));
    var newTxt = tmpTxt1 + '.' + tmpTxt2 + '.' + tmpTxt3;
    return newTxt;
  }
  
  // Move file to new folder
  function moveFile(fileId, newFolderId, oldFolderId) {
    var tempFile = DriveApp.getFileById(fileId);
    DriveApp.getFolderById(newFolderId).addFile(tempFile);
    DriveApp.getFolderById(oldFolderId).removeFile(tempFile);
    return;
  }
  
}
