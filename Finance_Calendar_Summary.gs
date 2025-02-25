var bigSpreadSheet;
var financeSS;
var summarySS;
var sheetRules;
var curYear = 0;
let headerRg = "A1:AU1";
let headerCells;
let yearRg = "A1"
let yearCell;
let calenderName;

let weekBordRg = [["C", "H"], ["I", "N"], ["O", "T"], ["U", "Z"], ["AA", "AF"], ["AG", "AL"], ["AM", "AR"]];
let weekDayRg = [["D", "G"], ["J", "M"], ["P", "S"], ["V", "Y"], ["AB", "AE"], ["AH", "AK"], ["AN", "AQ"]];

let sumBordRg = ["AT", "AU"];

let costCols = ["D", "J", "P", "V", "AB", "AH", "AN"];
let descCols = ["E", "K", "Q", "W", "AC", "AI", "AO"];
let expCols = ["F", "L", "R", "X", "AD", "AJ", "AP"];
let perCols = ["G", "M", "S", "Y", "AE", "AK", "AQ"];

let sumMonthCols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"];

const expenValueList = ["Food", "Rent", "Take-out", "Luxury", "Bills", "Necessities", "Medical", "Travel", "Car", "Saving", "Subscription", "Holiday", "Other"];
const personValueList = ["Person 1", "Person 2"];

const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

const formulaExpRef = "R[0]C[-1]";
const formulaExpCols = ["C[-5]", "C[-11]", "C[-17]", "C[-23]", "C[-29]", "C[-35]", "C[-41]"];
const formulaCostCols = ["C[-7]", "C[-13]", "C[-19]", "C[-25]", "C[-31]", "C[-37]", "C[-43]"];
const formulaFirstRows = ["R[1]", "R[5]"];
const formulaSecondRows = ["R[0]", "R[4]"];
const formulaThirdRows = ["R[-1]", "R[3]"];
const formulaFourthRows = ["R[-2]", "R[2]"];
const formulaFifthRows = ["R[-3]", "R[1]"];
const formulaTotalRows = ["R[-4]", "R[0]"];

const weekHeight = 10;
let weekInMonths = [];
let maxRowHeight = 2;
let lastMonthsRow = []; 

let firstWeekRows = [2, 11];
let firstDaysRow = 3;

const numFontFamily = "Lato";
const yearMonthFontSize = 36;
const weekDayFontSize = 24;
const dateFontSize = 14;
const genericTableFontSize = 10;
const numFormat = "£0.00";

const sumYearFontSize = 25;
const sumTitleFontSize = 20;
const sumExpBreakdownFontSize = 17;
const sumCategoryFontSize = 13;

function CreateCalenderAndSummary(year = 2024){
  calenderName = "Calendar " + year.toString();
  bigSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  financeSS = bigSpreadSheet.insertSheet(calenderName);
  financeSS.clearConditionalFormatRules(); //To ensure no overlap between sheets
  sheetRules = [];

  SetUp(year);

  for (let colsNum = 2; colsNum < 40; colsNum += 6){
    financeSS.setColumnWidths(colsNum, 2, 20);

  }

  financeSS.setColumnWidth(44, 20); //day boundaries
  financeSS.setColumnWidth(45, 28); //col separating days and weekly sums
  financeSS.setColumnWidth(1, 113); //Year col

  let costNumCol = 4;
  let descNumCol = 5;
  let expNumCol = 6;
  let perNumCol = 7;

  days.forEach(day => {
    financeSS.setColumnWidth(costNumCol, 60);
    financeSS.setColumnWidth(descNumCol, 90);
    financeSS.setColumnWidth(expNumCol, 85);
    financeSS.setColumnWidth(perNumCol, 65);

    costNumCol += 6;
    descNumCol += 6;
    expNumCol += 6;
    perNumCol += 6;

  })

  financeSS.setColumnWidth(46, 140); //Expense weekly sum col
  financeSS.setColumnWidth(47, 135); //Spent weekly sum col
  
  financeSS.setConditionalFormatRules(sheetRules);

  SpreadsheetApp.flush();

  SummarySheet(year);

  SpreadsheetApp.flush();

}

function SetUp(inYear){
  let cols = financeSS.getMaxColumns();

  if (cols != 49) {
    if (cols <= 26) {
      financeSS.insertColumnsAfter((cols - 1), (23 + (26 - cols)));

    } else {
      financeSS.deleteColumns(26, (financeSS.getMaxColumns() - 26));
      financeSS.insertColumnsAfter(25, 23);

    }
  }

  curYear = inYear;

  //Calculate the number of weeks in each month
  CalNoWeeks();

  //Working out how many rows are needed for the base calender
  let rows = financeSS.getMaxRows();

  for (let noWeeks = 0; noWeeks < weekInMonths.length; noWeeks++){
    maxRowHeight = maxRowHeight + (weekInMonths[noWeeks] * weekHeight);

  }

  if (rows < maxRowHeight) {
    financeSS.insertRowsAfter(rows, (maxRowHeight - rows));

  }

  //Flush the rows and column inserts so the sheet is updated in order for the rest of the formatting works
  SpreadsheetApp.flush();

  yearCell = financeSS.getRange(yearRg);
  yearCell.setValue(curYear);
  yearCell.setHorizontalAlignment("center");
  yearCell.setFontSize(yearMonthFontSize);

  headerCells = financeSS.getRange(headerRg);
  headerCells.setBorder(null, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  //Weekday Heading + Weekly Summary Heading
  WeekdayHeadings();

  MonthBorders();

  WeeklyBorders();

  DayTitles();

}

function WeekdayHeadings(){
  let weekDayCell;

  for (let weekDay = 0; weekDay < weekBordRg.length; weekDay++) {
    let dayRg = weekDayRg[weekDay][0] + "1:" + weekDayRg[weekDay][1] + "1";

    weekDayCell = financeSS.getRange(dayRg);
    weekDayCell.merge();
    weekDayCell.setValue(days[weekDay]);
    weekDayCell.setHorizontalAlignment("center");
    weekDayCell.setVerticalAlignment("middle");
    weekDayCell.setFontSize(weekDayFontSize);

  }

  let sumTitleCells = financeSS.getRange(sumBordRg[0] + "1:" + sumBordRg[1] + "1");
  sumTitleCells.merge();
  sumTitleCells.setValue("Weekly Summaries");
  sumTitleCells.setFontSize(weekDayFontSize);
  sumTitleCells.setHorizontalAlignment("center");

}

//Calculate the number of weeks in each month
function CalNoWeeks(){
  for (let noMonth = 0; noMonth < months.length; noMonth++){
    let firstDay = new Date(curYear, noMonth, 1); //First day of the current month
    let firstWeekDay = firstDay.getDay(); //Getting the first day of the current month
    let lastDay = new Date(curYear, noMonth + 1, 0); //Last day of the current month

    //A/N Alternate Solution: get the days on the month by using the final date of the month
    let deltaDays = Math.round((lastDay - firstDay) / (24 * 60 * 60 * 1000)); //Calculating how many days in the current month - 1 | Dates uses milliseconds for numerical operations

    //Days of the week start on Sunday in the Date class
    //Most months that start on a Sunday will always their last one or two days as a new week, this 7 is to account for that by adding an extra week in the following calculation
    //A/N Alternate Solution: As some places have Sunday as their first day of the week, make this more customisable by having a global with the first day of the week at index 0
    //                        using this global as the "firstWeekDay" in the following solution and getting the days of the month from the Date object rather than working it out.
    //                        Ultimately the same result, but would be cleaner and easier to understand.
    if (firstWeekDay == 0) {
      firstWeekDay = 7;

    }
  
    let weekNumber = Math.ceil((deltaDays + firstWeekDay)/ 7);

    weekInMonths.push(weekNumber);

  }
}

function MonthBorders(){
  let bordRgs = [];

  //Loop is incremented by 2 as each weekday border overlaps with the adjacent ones
  for (let weekDay = 0; weekDay < weekBordRg.length; weekDay += 2) {
    bordRgs.push(weekBordRg[weekDay][0] + "1:" + weekBordRg[weekDay][1] + (maxRowHeight - 1).toString());

  }

  bordRgs.push(sumBordRg[0] + "1:" + sumBordRg[1] + (maxRowHeight - 1).toString());

  bordCells = financeSS.getRangeList(bordRgs);
  bordCells.setBorder(null, true, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

}

function WeeklyBorders(){
  let sumWeeks = 0;

  //for each loop
  weekInMonths.forEach(week => {
    sumWeeks += week;

  });

  let weekRgs = [];
  let weekCells;

  //Doing alternative weeks as borderlines overlap
  //Start week 2, then week 4, but starts at 1, as indexes start at 0
  for(let week = 1; week < sumWeeks; week += 2){
    weekRgs.push(weekBordRg[0][0] + (firstWeekRows[0] + (weekHeight * week)).toString() + ":" + sumBordRg[1] + (firstWeekRows[1] + (weekHeight * week)).toString());

  }

  //As the previous loop is done for only odd numbers, the final borderline is not accounted for, therefore it is necessary to border it outside the loop
  weekRgs.push(weekBordRg[0][0] + (firstWeekRows[0] + (weekHeight * (sumWeeks - 1))).toString() + ":" + sumBordRg[1] + (firstWeekRows[0] + (weekHeight * sumWeeks) - 1).toString());
  weekCells = financeSS.getRangeList(weekRgs);
  weekCells.setBorder(true, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

}

function DayTitles(){
  let monthIndex;

  let firstDate;
  let firstDay;

  let lastDate;
  let lastDay;

  let curDay;

  let curRow;
  let dayRgs = [];
  let dayHeadingRgs = [];
  let dayTableRgs = [];
  let dayCells;

  let costRgs = [];
  let descRgs = [];
  let expRgs = [];
  let perRgs = [];
  let spentRgs = [];
  let totalRgs = [];
  let totalSumRgs = [];
  let fullCostRgs = [];

  let expDropDownRgs = [];
  let perDropDownRgs = [];

  let sumBordRgs = [];
  let sumExpRg;
  let sumSpentRg;
  let sumTableRgs = [];

  let fullSpentRgs = [[], [], [], [], []];

  let tableAllRgs;

  let nonDayCells;
  
  let weekCounter = 0;

  let previousRow = 0;

  months.forEach(month => {
    monthIndex = months.indexOf(month);
    curDay = 1;

    previousRow = MonthTitles(monthIndex, previousRow);

    firstDate = new Date(curYear, monthIndex, 1);
    firstDay = CorrectDayIndex(firstDate.getDay());

    curRow = (firstDaysRow + (weekHeight * weekCounter)).toString();
    dayRgs.push(weekBordRg[firstDay][0] + curRow + ":" + weekBordRg[6][1] + curRow);

    if (firstDay > 0){
      for (let nonDay = 0; nonDay < firstDay; nonDay++){
        nonDayCells = financeSS.getRange(weekDayRg[nonDay][0] + curRow + ":" + weekDayRg[nonDay][1] + ((parseInt(curRow) + 7)).toString());
        nonDayCells.merge();
        nonDayCells.setBackground("#efefef");

      }
    }

    curDay = DaysBordTitle(curRow, firstDay, 6, curDay);

    sumExpRg = sumBordRg[0] + curRow;
    sumSpentRg = sumBordRg[1] + curRow;

    expRgs.push(sumExpRg);
    spentRgs.push(sumSpentRg);
    sumBordRgs.push(sumExpRg + ":" + sumSpentRg);

    curRow = (parseInt(curRow) + 2).toString();
    dayHeadingRgs.push(weekBordRg[firstDay][0] + curRow + ":" + weekBordRg[6][1] + curRow);

    tableAllRgs = DaysBordTable(parseInt(curRow), firstDay, 6)

    dayTableRgs = dayTableRgs.concat(tableAllRgs.dayTableRgs);
    costRgs = costRgs.concat(tableAllRgs.costRgs);
    descRgs = descRgs.concat(tableAllRgs.descRgs);
    expRgs = expRgs.concat(tableAllRgs.expRgs);
    perRgs = perRgs.concat(tableAllRgs.perRgs);
    expDropDownRgs = expDropDownRgs.concat(tableAllRgs.expDropDownRgs);
    perDropDownRgs = perDropDownRgs.concat(tableAllRgs.perDropDownRgs);
    totalRgs.push(tableAllRgs.totalRg);
    totalSumRgs.push(tableAllRgs.totalSumRg);
    sumTableRgs.push(tableAllRgs.sumTableRg);
    fullCostRgs = fullCostRgs.concat(tableAllRgs.fullCostRgs);

    tableAllRgs.fullSpentRgs.forEach(rg => {
      fullSpentRgs[tableAllRgs.fullSpentRgs.indexOf(rg)].push(rg);

    })

    weekCounter++;

    for (let insideWeeks = 1; insideWeeks < (weekInMonths[months.indexOf(month)] - 1); insideWeeks++){
      curRow = (firstDaysRow + (weekHeight * weekCounter)).toString();
      dayRgs.push(weekBordRg[0][0] + curRow + ":" + weekBordRg[6][1] + curRow);

      curDay = DaysBordTitle(curRow, 0, 6, curDay);

      sumExpRg = sumBordRg[0] + curRow;
      sumSpentRg = sumBordRg[1] + curRow;

      expRgs.push(sumExpRg);
      spentRgs.push(sumSpentRg);
      sumBordRgs.push(sumExpRg + ":" + sumSpentRg);

      curRow = (parseInt(curRow) + 2).toString();
      dayHeadingRgs.push(weekBordRg[0][0] + curRow + ":" + weekBordRg[6][1] + curRow);

      tableAllRgs = DaysBordTable(parseInt(curRow), 0, 6)

      dayTableRgs = dayTableRgs.concat(tableAllRgs.dayTableRgs);
      costRgs = costRgs.concat(tableAllRgs.costRgs);
      descRgs = descRgs.concat(tableAllRgs.descRgs);
      expRgs = expRgs.concat(tableAllRgs.expRgs);
      perRgs = perRgs.concat(tableAllRgs.perRgs);
      expDropDownRgs = expDropDownRgs.concat(tableAllRgs.expDropDownRgs);
      perDropDownRgs = perDropDownRgs.concat(tableAllRgs.perDropDownRgs);
      totalRgs.push(tableAllRgs.totalRg);
      totalSumRgs.push(tableAllRgs.totalSumRg);
      sumTableRgs.push(tableAllRgs.sumTableRg);
      fullCostRgs = fullCostRgs.concat(tableAllRgs.fullCostRgs);

      tableAllRgs.fullSpentRgs.forEach(rg => {
        fullSpentRgs[tableAllRgs.fullSpentRgs.indexOf(rg)].push(rg);

      })

      weekCounter++;

    }

    lastDate = new Date(curYear, monthIndex + 1, 0);
    lastDay = CorrectDayIndex(lastDate.getDay());

    curRow = (firstDaysRow + (weekHeight * weekCounter)).toString();
    dayRgs.push(weekBordRg[0][0] + curRow + ":" + weekBordRg[lastDay][1] + curRow);

    if (lastDay < 6){
      for (let nonDay = lastDay + 1; nonDay <= 6; nonDay++){
        nonDayCells = financeSS.getRange(weekDayRg[nonDay][0] + curRow + ":" + weekDayRg[nonDay][1] + ((parseInt(curRow) + 7)).toString());
        nonDayCells.merge();
        nonDayCells.setBackground("#efefef");

      }
    }

    curDay = DaysBordTitle(curRow, 0, lastDay, curDay);

    sumExpRg = sumBordRg[0] + curRow;
    sumSpentRg = sumBordRg[1] + curRow;

    expRgs.push(sumExpRg);
    spentRgs.push(sumSpentRg);
    sumBordRgs.push(sumExpRg + ":" + sumSpentRg);

    curRow = (parseInt(curRow) + 2).toString();
    dayHeadingRgs.push(weekBordRg[0][0] + curRow + ":" + weekBordRg[lastDay][1] + curRow);

    tableAllRgs = DaysBordTable(parseInt(curRow), 0, lastDay);

    dayTableRgs = dayTableRgs.concat(tableAllRgs.dayTableRgs);
    costRgs = costRgs.concat(tableAllRgs.costRgs);
    descRgs = descRgs.concat(tableAllRgs.descRgs);
    expRgs = expRgs.concat(tableAllRgs.expRgs);
    perRgs = perRgs.concat(tableAllRgs.perRgs);
    expDropDownRgs = expDropDownRgs.concat(tableAllRgs.expDropDownRgs);
    perDropDownRgs = perDropDownRgs.concat(tableAllRgs.perDropDownRgs);
    totalRgs.push(tableAllRgs.totalRg);
    totalSumRgs.push(tableAllRgs.totalSumRg);
    sumTableRgs.push(tableAllRgs.sumTableRg);
    fullCostRgs = fullCostRgs.concat(tableAllRgs.fullCostRgs);

    tableAllRgs.fullSpentRgs.forEach(rg => {
      fullSpentRgs[tableAllRgs.fullSpentRgs.indexOf(rg)].push(rg);

    })

    weekCounter++;

  })

  dayCells = financeSS.getRangeList(dayRgs);
  dayCells.setBorder(true, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);

  dayCells = financeSS.getRangeList(dayHeadingRgs);
  dayCells.setFontWeight("bold");
  dayCells.setHorizontalAlignment("center");
  dayCells.setFontSize(genericTableFontSize);
  dayCells.setBorder(true, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);

  dayCells = financeSS.getRangeList(sumBordRgs);
  dayCells.setFontWeight("bold");
  dayCells.setHorizontalAlignment("center");
  dayCells.setVerticalAlignment("middle");
  dayCells.setFontSize(genericTableFontSize);
  dayCells.setBorder(true, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.DOTTED);

  dayCells = financeSS.getRangeList(sumTableRgs);
  dayCells.setFontSize(genericTableFontSize);
  dayCells.setBorder(true, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);

  dayCells = financeSS.getRangeList(dayTableRgs);
  dayCells.setBorder(null, true, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID);

  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground("#efefef")
    .setRanges(dayCells.getRanges())
    .build();

  sheetRules.push(rule);

  dayCells = financeSS.getRangeList(fullCostRgs);
  dayCells.setFontFamily(numFontFamily);
  dayCells.setHorizontalAlignment("right");
  dayCells.setNumberFormat(numFormat);

  dayCells = financeSS.getRangeList(costRgs);
  dayCells.setValue("Cost");

  dayCells = financeSS.getRangeList(descRgs);
  dayCells.setValue("Description");

  dayCells = financeSS.getRangeList(expRgs);
  dayCells.setValue("Expenses");

  dayCells = financeSS.getRangeList(perRgs);
  dayCells.setValue("Person");

  dayCells = financeSS.getRangeList(spentRgs);
  dayCells.setValue("Spent");

  dayCells = financeSS.getRangeList(totalRgs);
  dayCells.setFontSize(genericTableFontSize);
  dayCells.setHorizontalAlignment("center");
  dayCells.setValue("Total");

  CreateFormatRulesForDropDownMenus(expDropDownRgs, "Expense", financeSS);
  CreateFormatRulesForDropDownMenus(perDropDownRgs, "Person", financeSS);

  let formula;

  fullSpentRgs.forEach(spentRg => {
    dayCells = financeSS.getRangeList(spentRg);
    switch(fullSpentRgs.indexOf(spentRg)){
      // Formula: If the catergory cell is blank, output 0. Else, sum all the numbers in the previous week associated to the appopriate category.
      case 0:
        formula = "=if(isblank(" + formulaExpRef + "), 0, sum(sumif(" + formulaFirstRows[0] + formulaExpCols[0] + ":" + formulaFirstRows[1] + formulaExpCols[0] + ", " + formulaExpRef + ", " + formulaFirstRows[0] + formulaCostCols[0] + ":" + formulaFirstRows[1] + formulaCostCols[0] + "), sumif(" + formulaFirstRows[0] + formulaExpCols[1] + ":" + formulaFirstRows[1] + formulaExpCols[1] + ", " + formulaExpRef + ", " + formulaFirstRows[0] + formulaCostCols[1] + ":" + formulaFirstRows[1] + formulaCostCols[1] + "), sumif(" + formulaFirstRows[0] + formulaExpCols[2] + ":" + formulaFirstRows[1] + formulaExpCols[2] + ", " + formulaExpRef + ", " + formulaFirstRows[0] + formulaCostCols[2] + ":" + formulaFirstRows[1] + formulaCostCols[2] + "), sumif("    + formulaFirstRows[0] + formulaExpCols[3] + ":" + formulaFirstRows[1] + formulaExpCols[3] + ", " + formulaExpRef + ", " + formulaFirstRows[0] + formulaCostCols[3] + ":" + formulaFirstRows[1] + formulaCostCols[3] + "), sumif(" + formulaFirstRows[0] + formulaExpCols[4] + ":" + formulaFirstRows[1] + formulaExpCols[4] + ", " + formulaExpRef + ", " + formulaFirstRows[0] + formulaCostCols[4] + ":" + formulaFirstRows[1] + formulaCostCols[4] + "), sumif(" + formulaFirstRows[0] + formulaExpCols[5] + ":" + formulaFirstRows[1] + formulaExpCols[5] + ", " + formulaExpRef + ", " + formulaFirstRows[0] + formulaCostCols[5] + ":" + formulaFirstRows[1] + formulaCostCols[5] + "), sumif(" + formulaFirstRows[0] + formulaExpCols[6] + ":" + formulaFirstRows[1] + formulaExpCols[6] + ", " + formulaExpRef + ", " + formulaFirstRows[0] + formulaCostCols[6] + ":" + formulaFirstRows[1] + formulaCostCols[6] + ")))";
        break;

      case 1:
        formula = "=if(isblank(" + formulaExpRef + "), 0, sum(sumif(" + formulaSecondRows[0] + formulaExpCols[0] + ":" + formulaSecondRows[1] + formulaExpCols[0] + ", " + formulaExpRef + ", " + formulaSecondRows[0] + formulaCostCols[0] + ":" + formulaSecondRows[1] + formulaCostCols[0] + "), sumif(" + formulaSecondRows[0] + formulaExpCols[1] + ":" + formulaSecondRows[1] + formulaExpCols[1] + ", " + formulaExpRef + ", " + formulaSecondRows[0] + formulaCostCols[1] + ":" + formulaSecondRows[1] + formulaCostCols[1] + "), sumif(" + formulaSecondRows[0] + formulaExpCols[2] + ":" + formulaSecondRows[1] + formulaExpCols[2] + ", " + formulaExpRef + ", " + formulaSecondRows[0] + formulaCostCols[2] + ":" + formulaSecondRows[1] + formulaCostCols[2] + "), sumif("    + formulaSecondRows[0] + formulaExpCols[3] + ":" + formulaSecondRows[1] + formulaExpCols[3] + ", " + formulaExpRef + ", " + formulaSecondRows[0] + formulaCostCols[3] + ":" + formulaSecondRows[1] + formulaCostCols[3] + "), sumif(" + formulaSecondRows[0] + formulaExpCols[4] + ":" + formulaSecondRows[1] + formulaExpCols[4] + ", " + formulaExpRef + ", " + formulaSecondRows[0] + formulaCostCols[4] + ":" + formulaSecondRows[1] + formulaCostCols[4] + "), sumif(" + formulaSecondRows[0] + formulaExpCols[5] + ":" + formulaSecondRows[1] + formulaExpCols[5] + ", " + formulaExpRef + ", " + formulaSecondRows[0] + formulaCostCols[5] + ":" + formulaSecondRows[1] + formulaCostCols[5] + "), sumif(" + formulaSecondRows[0] + formulaExpCols[6] + ":" + formulaSecondRows[1] + formulaExpCols[6] + ", " + formulaExpRef + ", " + formulaSecondRows[0] + formulaCostCols[6] + ":" + formulaSecondRows[1] + formulaCostCols[6] + ")))";
        break;

      case 2:
        formula = "=if(isblank(" + formulaExpRef + "), 0, sum(sumif(" + formulaThirdRows[0] + formulaExpCols[0] + ":" + formulaThirdRows[1] + formulaExpCols[0] + ", " + formulaExpRef + ", " + formulaThirdRows[0] + formulaCostCols[0] + ":" + formulaThirdRows[1] + formulaCostCols[0] + "), sumif(" + formulaThirdRows[0] + formulaExpCols[1] + ":" + formulaThirdRows[1] + formulaExpCols[1] + ", " + formulaExpRef + ", " + formulaThirdRows[0] + formulaCostCols[1] + ":" + formulaThirdRows[1] + formulaCostCols[1] + "), sumif(" + formulaThirdRows[0] + formulaExpCols[2] + ":" + formulaThirdRows[1] + formulaExpCols[2] + ", " + formulaExpRef + ", " + formulaThirdRows[0] + formulaCostCols[2] + ":" + formulaThirdRows[1] + formulaCostCols[2] + "), sumif("    + formulaThirdRows[0] + formulaExpCols[3] + ":" + formulaThirdRows[1] + formulaExpCols[3] + ", " + formulaExpRef + ", " + formulaThirdRows[0] + formulaCostCols[3] + ":" + formulaThirdRows[1] + formulaCostCols[3] + "), sumif(" + formulaThirdRows[0] + formulaExpCols[4] + ":" + formulaThirdRows[1] + formulaExpCols[4] + ", " + formulaExpRef + ", " + formulaThirdRows[0] + formulaCostCols[4] + ":" + formulaThirdRows[1] + formulaCostCols[4] + "), sumif(" + formulaThirdRows[0] + formulaExpCols[5] + ":" + formulaThirdRows[1] + formulaExpCols[5] + ", " + formulaExpRef + ", " + formulaThirdRows[0] + formulaCostCols[5] + ":" + formulaThirdRows[1] + formulaCostCols[5] + "), sumif(" + formulaThirdRows[0] + formulaExpCols[6] + ":" + formulaThirdRows[1] + formulaExpCols[6] + ", " + formulaExpRef + ", " + formulaThirdRows[0] + formulaCostCols[6] + ":" + formulaThirdRows[1] + formulaCostCols[6] + ")))";
        break;

      case 3:
        formula = "=if(isblank(" + formulaExpRef + "), 0, sum(sumif(" + formulaFourthRows[0] + formulaExpCols[0] + ":" + formulaFourthRows[1] + formulaExpCols[0] + ", " + formulaExpRef + ", " + formulaFourthRows[0] + formulaCostCols[0] + ":" + formulaFourthRows[1] + formulaCostCols[0] + "), sumif(" + formulaFourthRows[0] + formulaExpCols[1] + ":" + formulaFourthRows[1] + formulaExpCols[1] + ", " + formulaExpRef + ", " + formulaFourthRows[0] + formulaCostCols[1] + ":" + formulaFourthRows[1] + formulaCostCols[1] + "), sumif(" + formulaFourthRows[0] + formulaExpCols[2] + ":" + formulaFourthRows[1] + formulaExpCols[2] + ", " + formulaExpRef + ", " + formulaFourthRows[0] + formulaCostCols[2] + ":" + formulaFourthRows[1] + formulaCostCols[2] + "), sumif("    + formulaFourthRows[0] + formulaExpCols[3] + ":" + formulaFourthRows[1] + formulaExpCols[3] + ", " + formulaExpRef + ", " + formulaFourthRows[0] + formulaCostCols[3] + ":" + formulaFourthRows[1] + formulaCostCols[3] + "), sumif(" + formulaFourthRows[0] + formulaExpCols[4] + ":" + formulaFourthRows[1] + formulaExpCols[4] + ", " + formulaExpRef + ", " + formulaFourthRows[0] + formulaCostCols[4] + ":" + formulaFourthRows[1] + formulaCostCols[4] + "), sumif(" + formulaFourthRows[0] + formulaExpCols[5] + ":" + formulaFourthRows[1] + formulaExpCols[5] + ", " + formulaExpRef + ", " + formulaFourthRows[0] + formulaCostCols[5] + ":" + formulaFourthRows[1] + formulaCostCols[5] + "), sumif(" + formulaFourthRows[0] + formulaExpCols[6] + ":" + formulaFourthRows[1] + formulaExpCols[6] + ", " + formulaExpRef + ", " + formulaFourthRows[0] + formulaCostCols[6] + ":" + formulaFourthRows[1] + formulaCostCols[6] + ")))";
        break;

      case 4:
        "=if(isblank(" + formulaExpRef + "), 0, sum(sumif(" + formulaFifthRows[0] + formulaExpCols[0] + ":" + formulaFifthRows[1] + formulaExpCols[0] + ", " + formulaExpRef + ", " + formulaFifthRows[0] + formulaCostCols[0] + ":" + formulaFifthRows[1] + formulaCostCols[0] + "), sumif(" + formulaFifthRows[0] + formulaExpCols[1] + ":" + formulaFifthRows[1] + formulaExpCols[1] + ", " + formulaExpRef + ", " + formulaFifthRows[0] + formulaCostCols[1] + ":" + formulaFifthRows[1] + formulaCostCols[1] + "), sumif(" + formulaFifthRows[0] + formulaExpCols[2] + ":" + formulaFifthRows[1] + formulaExpCols[2] + ", " + formulaExpRef + ", " + formulaFifthRows[0] + formulaCostCols[2] + ":" + formulaFifthRows[1] + formulaCostCols[2] + "), sumif("    + formulaFifthRows[0] + formulaExpCols[3] + ":" + formulaFifthRows[1] + formulaExpCols[3] + ", " + formulaExpRef + ", " + formulaFifthRows[0] + formulaCostCols[3] + ":" + formulaFifthRows[1] + formulaCostCols[3] + "), sumif(" + formulaFifthRows[0] + formulaExpCols[4] + ":" + formulaFifthRows[1] + formulaExpCols[4] + ", " + formulaExpRef + ", " + formulaFifthRows[0] + formulaCostCols[4] + ":" + formulaFifthRows[1] + formulaCostCols[4] + "), sumif(" + formulaFifthRows[0] + formulaExpCols[5] + ":" + formulaFifthRows[1] + formulaExpCols[5] + ", " + formulaExpRef + ", " + formulaFifthRows[0] + formulaCostCols[5] + ":" + formulaFifthRows[1] + formulaCostCols[5] + "), sumif(" + formulaFifthRows[0] + formulaExpCols[6] + ":" + formulaFifthRows[1] + formulaExpCols[6] + ", " + formulaExpRef + ", " + formulaFifthRows[0] + formulaCostCols[6] + ":" + formulaFifthRows[1] + formulaCostCols[6] + ")))";
        break;

    }

    dayCells.setNumberFormat(numFormat);
    dayCells.setFontFamily(numFontFamily);
    dayCells.setFormulaR1C1(formula);

  })

  // Formula: Sum all the numbers from the previous week
  formula = "=sum(" + formulaTotalRows[0] + formulaCostCols[0] + ":" + formulaTotalRows[1] + formulaCostCols[0] + ", " + formulaTotalRows[0] + formulaCostCols[1] + ":" + formulaTotalRows[1] + formulaCostCols[1] + ", " + formulaTotalRows[0] + formulaCostCols[2] + ":" + formulaTotalRows[1] + formulaCostCols[2] + ", " + formulaTotalRows[0] + formulaCostCols[3] + ":" + formulaTotalRows[1] + formulaCostCols[3] + ", " + formulaTotalRows[0] + formulaCostCols[4] + ":" + formulaTotalRows[1] + formulaCostCols[4] + ", " + formulaTotalRows[0] + formulaCostCols[5] + ":" + formulaTotalRows[1] + formulaCostCols[5] + ", " + formulaTotalRows[0] + formulaCostCols[6] + ":" + formulaTotalRows[1] + formulaCostCols[6] + ")";

  dayCells = financeSS.getRangeList(totalSumRgs);
  dayCells.setFontSize(genericTableFontSize);
  dayCells.setNumberFormat(numFormat);
  dayCells.setFontFamily(numFontFamily);
  dayCells.setFormulaR1C1(formula);

}

function CreateFormatRulesForDropDownMenus(rgs, option, sheet){

  let sheetCells = sheet.getRangeList(rgs);

  switch (option){
    case "Expense":
      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[0]) //Food
          .setBackground("#aaf7f4")
          .setFontColor("#03a9b9")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[1]) //Rent
          .setBackground("#ffcfc9")
          .setFontColor("#b10202")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[2]) //Take-out
          .setBackground("#ffc8aa")
          .setFontColor("#753800")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[3]) //Luxury
          .setBackground("#ffe5a0")
          .setFontColor("#473821")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[4]) //Bills
          .setBackground("#d4edbc")
          .setFontColor("#11734b")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[5]) //Necessities
          .setBackground("#bfe1f6")
          .setFontColor("#0a53a8")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[6]) //Medical
          .setBackground("#c6dbe1")
          .setFontColor("#215a6c")
          .setRanges(sheetCells.getRanges())
          .build());
          
      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[7]) //Travel
          .setBackground("#e6cff2")
          .setFontColor("#5a3286")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[8]) //Car
          .setBackground("#e6d488")
          .setFontColor("#ac9e64")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[9]) //Savings
          .setBackground("#f77acd")
          .setFontColor("#a7068e")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[10]) //Subscription
          .setBackground("#ff6161")
          .setFontColor("#690000")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[11]) //Holiday
          .setBackground("#b2b0ff")
          .setFontColor("#5551fb")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(expenValueList[12]) //Other
          .setBackground("#d3ffe0")
          .setFontColor("#006d21")
          .setRanges(sheetCells.getRanges())
          .build());
      
      break;

    case "Person":
      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(personValueList[0]) //Person 1
          .setBackground("#5a3286")
          .setFontColor("#e5cff2")
          .setRanges(sheetCells.getRanges())
          .build());
          
      sheetRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(personValueList[1]) //Person 2
          .setBackground("#473822")
          .setFontColor("#ffe5a0")
          .setRanges(sheetCells.getRanges())
          .build());

      sheetCells.setHorizontalAlignment("center");
          
      break;  
  }
}

function CorrectDayIndex(curIndex){
  if (curIndex == 0){
      curIndex = 6;

  } else {
      curIndex--;

  }

  return curIndex;

}

function DaysBordTitle(curRow, firstWeekBordRgIndex, lastWeekBordRgIndex, curDay){
  let dayRg;
  let dayCells;

  for (let days = firstWeekBordRgIndex; days <= lastWeekBordRgIndex; days++){
    dayRg = weekDayRg[days][0] + curRow + ":" + weekDayRg[days][1] + curRow;
    dayCells = financeSS.getRange(dayRg);
    dayCells.merge();
    dayCells.setFontSize(dateFontSize);
    dayCells.setHorizontalAlignment("center");
    dayCells.setValue(curDay.toString());
    dayCells.setFontWeight("bold");
    dayCells.setBorder(null, true, null, true, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);

    curDay++;
  }
  return curDay;

}

function CreateDropDownMenus(rgs, option, sheet){
  let dropDown;
  
  switch(option){
    case "Expense":
      dropDown = SpreadsheetApp.newDataValidation()
        .requireValueInList(expenValueList, true)
        .setAllowInvalid(false)
        .build();
      
      break;
    
    case "Person":
      dropDown = SpreadsheetApp.newDataValidation()
        .requireValueInList(personValueList, true)
        .setAllowInvalid(false)
        .build();
      break;
  }

  let dropDownCells = sheet.getRange(rgs);
  dropDownCells.setDataValidation(dropDown);

}

function DaysBordTable(curRow, firstDayBordRgIndex, lastDayBordRgIndex){
  let dayRgs = [];
  let costRgs = [];
  let descRgs = [];
  let expRgs = [];
  let perRgs = [];
  let sumTableRg;
  let totalRg;
  let totalSumRg;
  let fullCostRgs = [];

  let expDropDownRg;
  let perDropDownRg;

  let expDropDownRgs = [];
  let perDropDownRgs = [];

  let fullSpentRgs = [];

  let firstRow = (curRow + 1).toString();
  let lastRow = (curRow + 5).toString();

  for (let days = firstDayBordRgIndex; days <= lastDayBordRgIndex; days++){
    dayRgs.push(weekDayRg[days][0] + firstRow + ":" + weekDayRg[days][1] + lastRow);
    costRgs.push(costCols[days] + curRow.toString());
    descRgs.push(descCols[days] + curRow.toString());
    expRgs.push(expCols[days] + curRow.toString());
    perRgs.push(perCols[days] + curRow.toString());

    fullCostRgs.push(costCols[days] + firstRow + ":" + costCols[days] + lastRow);

    expDropDownRg = expCols[days] + firstRow + ":" + expCols[days] + lastRow;
    perDropDownRg = perCols[days] + firstRow + ":" + perCols[days] + lastRow;

    expDropDownRgs.push(expDropDownRg);
    perDropDownRgs.push(perDropDownRg);

    CreateDropDownMenus(expDropDownRg, "Expense", financeSS);
    CreateDropDownMenus(perDropDownRg, "Person", financeSS);

  }

  totalRg = sumBordRg[0] + lastRow;
  totalSumRg = sumBordRg[1] + lastRow;

  lastRow = (curRow + 4).toString();

  expDropDownRg = sumBordRg[0] + curRow.toString() + ":" + sumBordRg[0] + lastRow;
  expDropDownRgs.push(expDropDownRg);

  sumTableRg = sumBordRg[0] + curRow.toString() + ":" + sumBordRg[1] + lastRow;

  for (let i = 0; i <= 4; i++){
    fullSpentRgs[i] = sumBordRg[1] + (curRow + i).toString();

  }

  CreateDropDownMenus(expDropDownRg, "Expense", financeSS);

  return { 
    'dayTableRgs' : dayRgs,
    'costRgs' : costRgs,
    'descRgs' : descRgs,
    'expRgs' : expRgs,
    'perRgs' : perRgs,
    'expDropDownRgs' : expDropDownRgs,
    'perDropDownRgs' : perDropDownRgs,
    'totalRg' : totalRg,
    'totalSumRg' : totalSumRg,
    'sumTableRg' : sumTableRg,
    'fullCostRgs' : fullCostRgs,
    'fullSpentRgs' : fullSpentRgs

  };

}

function MonthTitles(monthIndex, previousRow){
  let monthRg;
  let monthTitleCells;
  let newPreviousRow = previousRow + (weekHeight * weekInMonths[monthIndex]);

  monthRg = "A" + (previousRow + 3).toString() + ":A" + newPreviousRow.toString();

  monthTitleCells = financeSS.getRange(monthRg);
  monthTitleCells.merge();
  monthTitleCells.setFontSize(yearMonthFontSize);
  monthTitleCells.setTextRotation(90);
  monthTitleCells.setVerticalAlignment("middle");
  monthTitleCells.setHorizontalAlignment("center");
  monthTitleCells.setValue(months[monthIndex]);

  //Need to Set colours

  if (monthIndex < 11) {
    monthRg = "A" + (newPreviousRow + 1).toString() + ":A" + (newPreviousRow + 2).toString();
    monthTitleCells = financeSS.getRange(monthRg);
    monthTitleCells.setBackground("#cccccc");

  }
  
  return newPreviousRow;

}

function SummarySheet(year){
  let summaryCells;
  let titleCol = ["A", "Q"];
  let dropDownCol = ["B", "P"];
  let formula;
  let monthRow = 2;

  summarySS = bigSpreadSheet.insertSheet("Summary " + year.toString());
  summarySS.clearConditionalFormatRules();
  sheetRules = [];

  yearCell = summarySS.getRange(yearRg);
  yearCell.setValue(year);
  yearCell.setFontSize(sumYearFontSize);
  yearCell.setHorizontalAlignment("center");

  summaryCells = summarySS.getRange(sumMonthCols[0] + "1:" + sumMonthCols[11] + "1");
  summaryCells.merge();
  summaryCells.setHorizontalAlignment("center");
  summaryCells.setFontSize(sumTitleFontSize);
  summaryCells.setValue("Monthly Summary");

  let monthIndex;
  let formulaRow;
  let formulaStartRow = 6;
  let formulaEndRow = 0;
  let formulaRef;
  let totalStartRow = 3;

  let p1TotalRow;
  let p2TotalRow;
  let p1IncomeRow;
  let p2IncomeRow
  let p1ExcessRow;
  let p2ExcessRow;

  sumMonthCols.forEach(monthCol => {
    monthIndex = sumMonthCols.indexOf(monthCol);
    summaryCells = summarySS.getRange(monthCol + monthRow.toString());
    summaryCells.setFontWeight("bold");
    summaryCells.setHorizontalAlignment("center");
    summaryCells.setFontSize(genericTableFontSize);
    summaryCells.setValue(months[monthIndex]);

    formulaRow = 3;
    if (monthIndex != 0) {
      formulaStartRow = formulaStartRow + (weekHeight * weekInMonths[monthIndex - 1]);

    }
    
    formulaEndRow = formulaEndRow + (weekHeight * weekInMonths[monthIndex]);

    expenValueList.forEach(expen => {
      formulaRef = dropDownCol[0] + formulaRow.toString()
      // Formula: If the catergory cell is blank, output 0. Else, sum all the numbers in the calender sheet associated to the appopriate category, excluding the weekly summary columns.
      formula = "if(isblank(" + formulaRef  + "), 0, sum(sumif('" + calenderName + "'!" + expCols[0] + formulaStartRow + ":" + expCols[0] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[0] + formulaStartRow + ":" + costCols[0] + formulaEndRow + "), sumif('" + calenderName + "'!" + expCols[1] + formulaStartRow + ":" + expCols[1] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[1] + formulaStartRow + ":" + costCols[1] + formulaEndRow + "), sumif('" + calenderName + "'!" + expCols[2] + formulaStartRow + ":" + expCols[2] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[2] + formulaStartRow + ":" + costCols[2] + formulaEndRow + "), sumif('" + calenderName + "'!" + expCols[3] + formulaStartRow + ":" + expCols[3] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[3] + formulaStartRow + ":" + costCols[3] + formulaEndRow + "), sumif('" + calenderName + "'!" + expCols[4] + formulaStartRow + ":" + expCols[4] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[4] + formulaStartRow + ":" + costCols[4] + formulaEndRow + "), sumif('" + calenderName + "'!" + expCols[5] + formulaStartRow + ":" + expCols[5] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[5] + formulaStartRow + ":" + costCols[5] + formulaEndRow + "), sumif('" + calenderName + "'!" + expCols[6] + formulaStartRow + ":" + expCols[6] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[6] + formulaStartRow + ":" + costCols[6] + formulaEndRow + ")))";

      summaryCells = summarySS.getRange(monthCol + formulaRow);
      summaryCells.setFormula(formula);
      summaryCells.setFontFamily(numFontFamily);
      summaryCells.setNumberFormat(numFormat);

      formulaRow++;
    })

    summaryCells.setBorder(null, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);

    formulaRow++;
    
    summaryCells = summarySS.getRange(monthCol + formulaRow.toString());
    summaryCells.setFormula("=sum(" + monthCol + totalStartRow.toString()  + ":" + monthCol + (totalStartRow + (expenValueList.length - 1)).toString() + ")");
    summaryCells.setHorizontalAlignment("right");
    summaryCells.setFontFamily(numFontFamily);
    summaryCells.setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);

    formulaRow++;
    p1TotalRow = formulaRow.toString();

    formulaRef = dropDownCol[0] + p1TotalRow;
    // Formula: If the catergory cell is blank, output 0. Else, sum all the numbers in the calender sheet associated to the appopriate person, excluding the weekly summary columns.
    formula = "if(isblank(" + formulaRef  + "), 0, sum(sumif('" + calenderName + "'!" + perCols[0] + formulaStartRow + ":" + perCols[0] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[0] + formulaStartRow + ":" + costCols[0] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[1] + formulaStartRow + ":" + perCols[1] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[1] + formulaStartRow + ":" + costCols[1] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[2] + formulaStartRow + ":" + perCols[2] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[2] + formulaStartRow + ":" + costCols[2] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[3] + formulaStartRow + ":" + perCols[3] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[3] + formulaStartRow + ":" + costCols[3] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[4] + formulaStartRow + ":" + perCols[4] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[4] + formulaStartRow + ":" + costCols[4] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[5] + formulaStartRow + ":" + perCols[5] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[5] + formulaStartRow + ":" + costCols[5] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[6] + formulaStartRow + ":" + perCols[6] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[6] + formulaStartRow + ":" + costCols[6] + formulaEndRow + ")))";

    summaryCells = summarySS.getRange(monthCol + formulaRow);
    summaryCells.setFormula(formula);
    summaryCells.setHorizontalAlignment("right");
    summaryCells.setFontFamily(numFontFamily);
    summaryCells.setNumberFormat(numFormat);

    formulaRow++;
    p2TotalRow = formulaRow.toString();

    formulaRef = dropDownCol[0] + p2TotalRow;
    // Formula: If the catergory cell is blank, output 0. Else, sum all the numbers in the calender sheet associated to the appopriate person, excluding the weekly summary columns.
    formula = "if(isblank(" + formulaRef  + "), 0, sum(sumif('" + calenderName + "'!" + perCols[0] + formulaStartRow + ":" + perCols[0] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[0] + formulaStartRow + ":" + costCols[0] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[1] + formulaStartRow + ":" + perCols[1] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[1] + formulaStartRow + ":" + costCols[1] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[2] + formulaStartRow + ":" + perCols[2] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[2] + formulaStartRow + ":" + costCols[2] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[3] + formulaStartRow + ":" + perCols[3] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[3] + formulaStartRow + ":" + costCols[3] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[4] + formulaStartRow + ":" + perCols[4] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[4] + formulaStartRow + ":" + costCols[4] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[5] + formulaStartRow + ":" + perCols[5] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[5] + formulaStartRow + ":" + costCols[5] + formulaEndRow + "), sumif('" + calenderName + "'!" + perCols[6] + formulaStartRow + ":" + perCols[6] + formulaEndRow + ", " + formulaRef + ", '" + calenderName + "'!" + costCols[6] + formulaStartRow + ":" + costCols[6] + formulaEndRow + ")))";

    summaryCells = summarySS.getRange(monthCol + formulaRow);
    summaryCells.setFormula(formula);
    summaryCells.setHorizontalAlignment("right");
    summaryCells.setFontFamily(numFontFamily);
    summaryCells.setNumberFormat(numFormat);

    formulaRow += 2;
    p1IncomeRow = formulaRow.toString();
    p2IncomeRow = (formulaRow + 1).toString();

    summaryCells = summarySS.getRange(monthCol + p1IncomeRow + ":" + monthCol + p2IncomeRow);
    summaryCells.setNumberFormat(numFormat);
    summaryCells.setHorizontalAlignment("right");
    summaryCells.setFontFamily(numFontFamily);
    summaryCells.setValue(0);

    formulaRow += 3;
    p1ExcessRow = formulaRow.toString();

    summaryCells = summarySS.getRange(monthCol + p1ExcessRow);
    summaryCells.setNumberFormat(numFormat);
    summaryCells.setHorizontalAlignment("right");
    summaryCells.setFontFamily(numFontFamily);
    summaryCells.setFormula("=" + monthCol + p1IncomeRow  + " - " + monthCol + p1TotalRow);

    formulaRow++;
    p2ExcessRow = formulaRow.toString();

    summaryCells = summarySS.getRange(monthCol + p2ExcessRow);
    summaryCells.setNumberFormat(numFormat);
    summaryCells.setHorizontalAlignment("right");
    summaryCells.setFontFamily(numFontFamily);
    summaryCells.setFormula("= " + monthCol + p2IncomeRow + " - " + monthCol + p2TotalRow);

    formulaRow++;

    summaryCells = summarySS.getRange(monthCol + formulaRow.toString());
    summaryCells.setNumberFormat(numFormat);
    summaryCells.setHorizontalAlignment("right");
    summaryCells.setFontFamily(numFontFamily);
    summaryCells.setFormula("= " + monthCol + p1ExcessRow  + " + " + monthCol + p2ExcessRow);
     

  })

  let yearTotalCol = "O";

  summaryCells = summarySS.getRange(yearTotalCol + monthRow.toString());
  summaryCells.setValue("Year Total");
  summaryCells.setFontWeight("bold");
  summaryCells.setFontSize(genericTableFontSize);

  formula = "=SUM(R[0]C[-12]:R[0]C[-1])";

  let yearTotalRgs = [];
  let firstRow = monthRow + 1
  let lastRow = monthRow + expenValueList.length;

  yearTotalRgs.push(yearTotalCol + (firstRow).toString() + ":" + yearTotalCol + (lastRow).toString());

  firstRow = lastRow + 2;
  lastRow = firstRow + 2;

  yearTotalRgs.push(yearTotalCol + (firstRow).toString() + ":" + yearTotalCol + (lastRow).toString());

  firstRow = lastRow + 2;
  lastRow = firstRow + 1;

  yearTotalRgs.push(yearTotalCol + (firstRow).toString() + ":" + yearTotalCol + (lastRow).toString());

  firstRow = lastRow + 2;
  lastRow = firstRow + 3;

  yearTotalRgs.push(yearTotalCol + (firstRow).toString() + ":" + yearTotalCol + (lastRow).toString());

  summaryCells = summarySS.getRangeList(yearTotalRgs);
  summaryCells.setNumberFormat(numFormat);
  summaryCells.setFormulaR1C1(formula);
  summaryCells.setFontFamily(numFontFamily);

  CreateCategoryTitleCols(monthRow, dropDownCol[0] , titleCol[0], "Left");
  CreateCategoryTitleCols(monthRow, dropDownCol[1] , titleCol[1], "Right");

}

function CreateCategoryTitleCols(row, dropDownCol, titleCol, pos){

  let curRow = row;
  let topExpRow = 3;
  let dropDownRg; 

  summaryCells = summarySS.getRange(dropDownCol + curRow.toString());
  summaryCells.setValue("Categories");
  summaryCells.setHorizontalAlignment("center");
  summaryCells.setFontSize(genericTableFontSize);
  summaryCells.setFontWeight("bold");

  curRow += expenValueList.length;

  summaryCells = summarySS.getRange(titleCol + topExpRow.toString() + ":" + titleCol + curRow.toString());
  summaryCells.merge();
  summaryCells.setFontSize(sumExpBreakdownFontSize);
  summaryCells.setValue("Expense Breakdown");
  summaryCells.setHorizontalAlignment("center");
  summaryCells.setVerticalAlignment("middle");

  if (pos == "Left") {
    summaryCells.setTextRotation(90);
    
  } else {
    summaryCells.setTextRotation(-90);

  }
  

  dropDownRg = dropDownCol + topExpRow.toString() + ":" + dropDownCol + curRow.toString();

  CreateDropDownMenus(dropDownRg, "Expense", summarySS);
  CreateFormatRulesForDropDownMenus([dropDownRg], "Expense", summarySS);

  expenValueList.forEach(expenseName =>{
    summaryCells = summarySS.getRange(dropDownCol + (topExpRow + expenValueList.indexOf(expenseName)).toString());
    summaryCells.setFontSize(genericTableFontSize);
    summaryCells.setValue(expenseName);

  })

  curRow += 2;

  summaryCells = summarySS.getRange(titleCol + curRow.toString() + ":" + titleCol + (curRow + 2).toString());
  summaryCells.merge();
  summaryCells.setFontSize(sumCategoryFontSize);
  summaryCells.setHorizontalAlignment("center");
  summaryCells.setVerticalAlignment("middle");
  summaryCells.setValue("Totals");

  summaryCells = summarySS.getRange(dropDownCol + curRow.toString());
  summaryCells.setValue("Combined");
  summaryCells.setHorizontalAlignment("center");
  summaryCells.setFontSize(genericTableFontSize);

  curRow++;

  curRow = SumSetPersons(curRow, dropDownCol);

  curRow += 2;

  summaryCells = summarySS.getRange(titleCol + curRow.toString() + ":" + titleCol + (curRow + 1).toString());
  summaryCells.merge();
  summaryCells.setFontSize(sumCategoryFontSize);
  summaryCells.setHorizontalAlignment("center");
  summaryCells.setVerticalAlignment("middle");
  summaryCells.setValue("Income");

  curRow = SumSetPersons(curRow, dropDownCol);

  curRow += 2;

  summaryCells = summarySS.getRange(titleCol + curRow.toString() + ":" + titleCol + (curRow + 2).toString());
  summaryCells.merge();
  summaryCells.setFontSize(sumCategoryFontSize);
  summaryCells.setHorizontalAlignment("center");
  summaryCells.setVerticalAlignment("middle");
  summaryCells.setValue("Excess");

  curRow = SumSetPersons(curRow, dropDownCol);

  curRow++;

  summaryCells = summarySS.getRange(dropDownCol + curRow.toString());
  summaryCells.setFontSize(genericTableFontSize);
  summaryCells.setHorizontalAlignment("center");
  summaryCells.setValue("Combined");

  summarySS.setConditionalFormatRules(sheetRules);


}

function SumSetPersons(curRow, dropDownCol){
  let summaryCells;
  let dropDownRg = dropDownCol + curRow.toString() + ":" + dropDownCol + (curRow + 1).toString();

  CreateDropDownMenus(dropDownRg, "Person", summarySS);
  CreateFormatRulesForDropDownMenus([dropDownRg], "Person", summarySS);
  summaryCells = summarySS.getRange(dropDownCol + curRow.toString());
  summaryCells.setValue(personValueList[0]);

  curRow++;

  summaryCells = summarySS.getRange(dropDownCol + curRow.toString());
  summaryCells.setValue(personValueList[1]);

  return curRow;

}
