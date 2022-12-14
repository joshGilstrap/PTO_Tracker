/** @OnlyCurrentDoc
 * Run on new details reports. Updates each employee that is in employeeArr
 * with new Approved or Cancel Approved. Handles everything from copy/pasting
 * calendars between sheets to updating individual cells with dates off.
 * @customfunction
 */
function updateMainSheet() {
  makeSafetyCopy();
  makeEmployeeObjectArray();

  let monthArr = makeMonthsArr();
  let monthDisplayArr = [];
  for(let i = 0; i < monthArr.length; ++i) {
    monthDisplayArr.push(monthArr[i].getDisplayValues());
  }

  spreadsheet.setNamedRange("temprange", mainSheet.getRange(1, 1, fullEmployeeChart.getHeight(), fullEmployeeChart.getWidth()));
  spreadsheet.setNamedRange("temprange2", updateSheet.getRange(1,1,fullEmployeeChart.getHeight(),fullEmployeeChart.getWidth()));
  let namedRanges = spreadsheet.getNamedRanges();
  namedRanges.sort(compare);
  let tempRange = namedRanges[namedRanges.length - 2];
  let tempRange2 = namedRanges[namedRanges.length - 1];

  let copyHere = updateSheet.getRange(1,1);

  for(let i = 0; i < employeeArr.length; ++i) {
    let employeeRow = (employeeArr[i].employeePos * 10) + 1;

    tempRange.setRange(mainSheet.getRange(employeeRow, 1, fullEmployeeChart.getHeight(), fullEmployeeChart.getWidth()));
    tempRange.getRange().copyTo(copyHere);

    traverseDates(employeeArr[i].dateObjArr, monthArr, monthDisplayArr);

    tempRange2.setRange(updateSheet.getRange(1,1,fullEmployeeChart.getHeight(),fullEmployeeChart.getWidth()));
    tempRange2.getRange().copyTo(tempRange.getRange());
    fullEmployeeChart.copyTo(copyHere);
  }
  spreadsheet.removeNamedRange("temprange");
  spreadsheet.removeNamedRange("temprange2");
  return;
}


/** @OnlyCurrentDoc
 * Glorified for loop, outer loop to the inner loops of processDateArray()
 */
function traverseDates(objArr, monthArr, displayArr) {
  for(let i = 0; i < objArr.length; ++i) {
    processDateArray(objArr[i], objArr[i].dateArr, monthArr, displayArr);
  }
}


/** @OnlyCurrentDoc
 * Iterates through the dateArr in DateObj array and finds the
 * corresponding cell in the calendar.
 */
function processDateArray(objArr, arr, months, display) {
  for(let i = 0; i < arr.length; ++i) {
    let monthValues = display[arr[i][1] - 1];
    for(let j = 0; j < monthValues.length; ++j) {
      let weekRange = monthValues[j];
      let doneWithDate = false;
      for(let k = 0; k < weekRange.length; ++k) {
        if(parseInt(weekRange[k]) === arr[i][2]) {
          var cellC = months[arr[i][1] - 1].getColumn() + k;
          var cellR = months[arr[i][1] - 1].getRow() + j;
          var cell = updateSheet.getRange(cellR,cellC);

          handleCellManip(objArr, cell, i);

          doneWithDate = true;
          break;
        }
      }
      if(doneWithDate) {
        break;
      }
    }
  }
}


/** @OnlyCurrentDoc
 * Updates calendar cell with appropriate information.
 */
function handleCellManip(obj, range, place) {
  if(obj.status === 'Cancel Approved') {
    range.setBackground('white');
    range.setNote(null);
  }
  else {
    if(obj.type === 'Vacation') {
      range.setBackground('#8eaadc');
    }
    else {
      range.setBackground('#f4b082');
    }
    range.setNote(
      `\n${obj.type}: ${obj.dateArr[place][3]}\n
      Requested Date : ${obj.dateArr[place][4]}\n
      Approved by: ${obj.dateArr[place][5]}`
      );
  }
}


/** @OnlyCurrentDoc
 * Make an array of all calendar months ranges with respect to
 * UpdateSheet's layout.
 * @return arr - array of months ranges
 */
function makeMonthsArr() {
  let arr = [];
  let updateRanges = updateSheet.getNamedRanges();
  updateRanges.sort(compare);
  for(var i = 0; i < updateRanges.length; ++i) {
    if(updateRanges[i].getName().indexOf("Month") > -1) {
      arr.push(updateRanges[i].getRange());
    }
  }
  return arr;
}



