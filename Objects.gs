/**
 * Employee objects are used to hold all info about employee and all time off
 * requests for them on Details. Used for calendar updates
 * 
 * name - employee name
 * hireDate - date of hire
 * vacaHours - current vacation time bank
 * sickHours - current sick time bank
 * dateObjArr - array of date objects, each element is one instance from Details
 */
function EmployeeObj() {
  this.name = '';
  this.hireDate = '';
  this.vacaHours = 0;
  this.sickHours = 0;
  this.dateObjArr = [];
  this.employeePos = -1;
};

/**
 * Date objects are used to grab individual instances of time off from the Details sheet
 * 
 * dateString - raw name from Details (e.g. '10/17/2022', '12/12/2022 - 12/15/2022')
 * dateArr - holds all info about each particualr day of that request
 * days - how may calendar days affected
 * type - 'Vacation' or 'Sick'
 * status - 'Approved' or 'Cancel Approved'
 * 
 * Date array consists of 6 pieces of information in this particular order.
 * [0] = date requested (string), [1] = month (int), [2] = day (int),
 * [3] = hours taken (int), [4] = date of request (string), [5] = approver (string),
 * [6] = is a holiday (bool)
 */
function DateObj() {
  this.dateString = '';
  this.dateArr = [];
  this.days = 0;
  this.type = '';
  this.status = '';
};







