/** buildEmployeeInfo.gs
 * Creates an array of employee objects, each containing all time off
 * requests labeled 'Approved' or 'Cancel Approved' on 'Details' sheet
 */

//Global employee array
var employeeArr = [];


/** @OnlyCurrentDoc
 * Builds an array of employee objects for processing. Employee
 * objects contain all time off info from Details sheet per employee.
 * @customfunction
 */
function makeEmployeeObjectArray() {
  let reportArr = grabReportNames();
  let directoryArr = grabDirectoryNames();

  for(let i = 0; i < reportArr.length; ++i) {
    if(!directoryArr.includes(reportArr[i])) {continue;}

    let tempEmp = new EmployeeObj();
    tempEmp.name = reportArr[i];
    tempEmp.employeePos = directoryArr.indexOf(tempEmp.name);
    tempEmp.hireDate = directoryHireDate[directoryArr.indexOf(tempEmp.name) + 1][0];

    tempEmp.vacaHours = findPTOHours(tempEmp.name, "Vacation");
    tempEmp.sickHours = findPTOHours(tempEmp.name, "Sick");

    createDateObjArr(tempEmp);
    employeeArr.push(tempEmp);
  }
}


/** @OnlyCurrentDoc
 * Operates the same way findPTO() does
 * @param targetName - employee name
 * @param typePTO - string literal "Vacation" or "Sick"
 * @return - number of hours, '--' for N/A
 * @customfunction
 */
function findPTOHours(targetName, typePTO) {
  let lastRow = directoryN.getLastRow();
  for(let i = 0; i < lastRow; ++i) {
    if(targetName == breakdownNames[i][0]) {
      let rowNum = i;
      let typeName = breakdownType[rowNum][0];
      if(typePTO != typeName) {
        typeName = breakdownType[rowNum + 1][0];
        if(typePTO != typeName) {return '--';}
        return breakdownBalance[rowNum + 1][0];
      }
      return breakdownBalance[rowNum][0];
    }
  }
}


/** @OnlyCurrentDoc
 * Scans Details sheet for employee's name and timeoff status. Generates and
 * pushes a date object into the employee's dateObjArr for each day out,
 * including ranges.
 * @param emp - Employee object
 */
function createDateObjArr(emp) {
  let lastRow = reportSheet.getLastRow();
  for(let i = 0; i < lastRow; ++i) {
    if(emp.name === reportNames[i][0]) {
      if(reportStatus[i][0] === 'Approved' || reportStatus[i][0] === 'Cancel Approved') {
        let tempDate = new DateObj();
        tempDate.dateString = reportDates[i][0];
        tempDate.status = reportStatus[i][0];
        tempDate.type = reportPolicy[i][0];
        if(reportDates[i][0].length > 10) {
          tempDate = makeDateRangeArr(tempDate, i);
        }
        else {
          tempDate.dateArr.push(makeDateArr(tempDate.dateString, i));
        }

        if(tempDate.status === 'Cancel Approved') {
          emp.dateObjArr.unshift(tempDate);
        }
        else {
          emp.dateObjArr.push(tempDate);
        }
      }
    }
  }
}


/** @OnlyCurrentDoc
 * Breaks a range of dates off up into individual dates, pushes all into
 * dateObj.
 * @param dateObj - Date object
 * @param place - Current row on Details sheet
 * @return dateObj - The modified date object passed in
 */
function makeDateRangeArr(dateObj, place) {
  let startMonth = parseInt(dateObj.dateString.substring(0,2));
  let startDay = parseInt(dateObj.dateString.substring(3,5));
  let endMonth = parseInt(dateObj.dateString.substring(13,15));
  let endDay = parseInt(dateObj.dateString.substring(16,18));

  if(startMonth !== endMonth) {
    let daysLeftInMonth = determineNumDays(startMonth) - startDay + 1;
    for(let i = 0; i < daysLeftInMonth; ++i) {
      let string = `${startMonth}/${startDay + i}/${CURRENT_YEAR}`;
      if(startMonth < 10) {string += '0';}
      dateObj.dateArr.push(makeDateArr(string, place));
    }

    for(let i = 1; i <= endDay; ++i) {
      let string;
      if(i < 10) {
        string = `${endMonth}/0${i}/${CURRENT_YEAR}`;
      }
      else {
        string = `${endMonth}/${i}/${CURRENT_YEAR}`;
      }
      dateObj.dateArr.push(makeDateArr(string, place));
    }
  }
  else {
    for(let i = 0; i < (endDay - startDay) + 1; ++i) {
      let string;
      if(startMonth < 10) {
        string = `0${startMonth}/${startDay + i}/${CURRENT_YEAR}`;
      }
      else {
        string = `${startMonth}/${startDay + i}/${CURRENT_YEAR}`;
      }
      dateObj.dateArr.push(makeDateArr(string, place));
    }
  }
  return dateObj;
}


/** @OnlyCurrentDoc
 * Create the date array.
 * @param string - full date in mm/dd/yyyy format
 * @param place - current row on Details sheet
 * @return arr - full date array
 * 
 * Date array consists of 6 pieces of information in this particular order.
 * [0] = date requested (string), [1] = month (int), [2] = day (int),
 * [3] = hours taken (int?), [4] = date of request (string), [5] = approver (string)
 * [6] = is a holiday (bool)
 */
function makeDateArr(string, place) {
  let arr = [];
  arr.push(string);
  arr.push(parseInt(string.substring(0,2)));
  arr.push(parseInt(string.substring(3,5)));
  arr.push(parseInt(reportHours[place][0]));
  arr.push(reportSubmittedDate[place][0]);
  arr.push(reportReviewedBy[place][0]);
  arr.push(isHoliday(string));
  return arr;
}


/** @OnlyCurrentDoc
 * Grab all employee names that have time requests on Details sheet
 * @return arr - array of time off names
 */
function grabReportNames() {
  let arr = [];
  let lastRow = reportSheet.getLastRow();
  for(let i = 1; i < lastRow; ++i) {
    let currentName = reportNames[i][0];
    if(!arr.includes(currentName)) {
      arr.push(currentName);
    }
  }
  return arr;
}


/** @OnlyCurrentDoc
 * Make an array of all employee names in string form
 * @return arr - array of directory names
 */
function grabDirectoryNames() {
  let arr = [];
  for(let i = 1; i < directorySheet.getLastRow(); ++i) {
    arr.push(directoryNames[i][0]);
  }
  return arr;
}


/** @OnlyCurrentDoc
 * Useful for determining how many months are in the month selected
 * @param month - month in integer format
 * @return - the number of days in that month - 0 indexed
 */
function determineNumDays(month) {
  let monthEndDays = [31,28,31,30,31,30,31,31,30,31,30,31];
  return monthEndDays[month - 1];
}


/** @OnlyCurrentDoc
 * Determines if a date worked was an observed holiday
 * WILL NEED TO BE HARD-CODED ANNUALLY
 * @return - true/false holiday status
 */
function isHoliday(date) {
  let holidays2022 = [`02/21/${CURRENT_YEAR}`,`05/30/${CURRENT_YEAR}`,
    `07/04/${CURRENT_YEAR}`,`09/04/${CURRENT_YEAR}`,
    `11/23/${CURRENT_YEAR}`,`12/25/${CURRENT_YEAR}`];
  return holidays2022.includes(date);
}


function compare(a,b) {
  if(a.getName() > b.getName()) {return 1;}
  else if(a.getName() < b.getName()) {return -1;}

  return 0;
}















