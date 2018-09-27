/**
Author: Darrick Ross
Date: 9/6/18

Description: Custom Google Spreadsheet function that adds a new function to be used in google
spreadsheet that will calcuated the average of an array of items and is capable of dropping the
smallest 'n' number of items. This was mainly built for grade calculation.
*/


/**
Needed a function that would add in order to use reduce to sum up an array
This is not supposed to be found in the autocomplete and is for internal use
*/
function add_function_helper(a, b) {
    return a + b;
}


/**
 * Calculate the average of an array and drop the smallest n number of values
 *
 * @param {Array} Array of data, {number} dropNum the number of data point to exclude from average
 * @return The Average of an array, excluding the smallest n number of elements
 * @customfunction
 */
function avgWithDroppedItems(range, drop) {
  //obtain the array from the spreadsheet at the range passed in
  var r = SpreadsheetApp.getActiveSpreadsheet().getRange(range);

  var dropNum = Number(drop)
  if (isNaN(dropNum)) {
    try {
      var dropNumRange = SpreadsheetApp.getActiveSpreadsheet().getRange(drop);
      var dropRangeArray = r.getValues();
      dropNum = Number(dropRangeArray[0]);
    }
    catch (e) {
      return "2nd is NaN";
    }
  }


  //Convert the values into an array
  var dataArray = []
  dataArray = r.getValues();

  //If the user asks to drop 0 or less just return a standard avg
  if (dropNum < 1) {
    return dataArray/dataArray.length;
  }

  //If the user wants to drop greater than or equal to then the number of elements in the array
  //Then return 0, because 'technically' the math checks out...
  if (dropNum >= dataArray.length) {
    return 0;
  }

  //Now we know that the user want to really calc an avg with dropped data points(s) and some data remaining

  //Make a sorted array from least to greatest, then splice off the first 'dropNum'
  var sortedArray = dataArray.slice(0).sort().splice(dropNum);

  var sortedSum = 0;

  //sum up array
  for (var i = 0; i < sortedArray.length; i++) {
    sortedSum += Number(sortedArray[i]);
  }

  //return adjusted avg
  return sortedSum/sortedArray.length;
}
