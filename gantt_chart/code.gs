// Custom Function
// https://developers.google.com/apps-script/guides/sheets/functions

// Event
// https://developers.google.com/apps-script/guides/triggers/events

// public sheet with scripts: https://docs.google.com/spreadsheets/d/1nEUos1LPKZsRsHTEdk9JvyiO4u_qqGgHj3z77WmI-Vg/edit?usp=sharing
// Check output on https://script.google.com/u/0/home/projects/

/*
*　Gantt chart： Fill the start and end date, and scripts automatically fill the cell with the length of date.
*/

function getThisCellDate(begin_date, nth_column){
  const day_start_column = 4; // AD
  const diff_days = (nth_column - day_start_column)
  const myDate = Moment.moment(begin_date).add(diff_days, 'd')
  return myDate;
}

function myGetDay(begin_date, nth_column){
  const myDate = getThisCellDate(begin_date, nth_column)
  return myDate.date();
}

function onEdit(e){
  const HEX_COLOR = {
    FILL: "#a6d4fa",
    ERASE: "#FFF",
    VALID_RANGE: "#81c784",
    INVALID_RANGE: "#f6a5c0",
  };
  var day_start_column = 4; // AD
  var day_start_column_from_begin_date = 3; // AD
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var range_of_start_date = sheet.getRange("A4:A");
  for(var i=1; true; i++){
    var cell = range_of_start_date.getCell(i, 1);
    if (cell.getValue() == '') break;
    else{
      this_start_date = Moment.moment(cell.getValue());
      begin_date = Moment.moment(sheet.getRange("B2").getValue());
      var diff_days = this_start_date.diff(begin_date, 'days');
      if(diff_days < 0) break;
      
      var first_cell_to_fill = day_start_column_from_begin_date + diff_days
      var date_range = cell.offset(0, 2);      
      var date_range_num = date_range.getValue();
      
      if(date_range_num > 0) date_range.setBackground(HEX_COLOR.VALID_RANGE);
      else date_range.setBackground(HEX_COLOR.INVALID_RANGE);
      
      sheet.getRange(cell.getRow(), day_start_column, 1, sheet.getLastColumn() - day_start_column + 1).setBackground(HEX_COLOR.ERASE);
      for(var day=0; day<date_range_num; day++ ){
        cell.offset(0, first_cell_to_fill+day).setBackground(HEX_COLOR.FILL)
      }
    }
  } 
}
