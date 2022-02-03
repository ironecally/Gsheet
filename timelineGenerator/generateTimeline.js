
function generateTimeline() {
  // set time range
  var start_date = new Date("2022-01-01 00:00:00");
  var end_date = new Date("2022-03-01 00:00:00");

  const default_sheet_col = ["Task", "PIC", "Working Days", "Status", "{calendar}"]
  const default_sheet_width = [200, 100, 50, 70, 30]

  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];

  var activeSheet = SpreadsheetApp.getActiveSheet()
  activeSheet.clear()

  //get sheet name
  Logger.log("modifying: "+activeSheet.getName());

  //generate merged column based on value from default_sheet_col (without calendar)
  var start_col = 65;
  for(let i=0; i< default_sheet_col.length;i++){
    if(is_calendar(default_sheet_col[i])){
      start_col++;
      continue;
    }
    else{
      let current_column = String.fromCharCode(start_col);
      let range_idx =  current_column+"1:"+current_column+"2";
      activeSheet.setColumnWidth(i+1,default_sheet_width[i]);
      activeSheet.getRange(range_idx).merge().setValue(default_sheet_col[i]);
    
      start_col++;
    }
  }

  //froze the view for rows & columns
  activeSheet.setFrozenColumns(default_sheet_col.length-1);
  activeSheet.setFrozenRows(2);

  //generate column as needed by calendar
  var date_list = getDates(start_date,end_date);
  let calendar_column = String.fromCharCode(65+default_sheet_col.findIndex(is_calendar));
  activeSheet.insertColumnsAfter(default_sheet_col.findIndex(is_calendar),date_list.length);
  
  //generate date column
  let col_index = default_sheet_col.findIndex(is_calendar)+1;
  for(let i=0; i<date_list.length; i++){
    activeSheet.getRange(2,col_index).setValue(date_list[i].getDate());
    activeSheet.setColumnWidth(col_index,default_sheet_width[4]);
    col_index++;
  }

  //generate weekend
  col_index = default_sheet_col.findIndex(is_calendar)+1;
  let curr_range;
  for(let i=0; i<date_list.length; i++){
    
    if(date_list[i].getDay() == 0 || date_list[i].getDay() == 6){
      curr_range = activeSheet.getRange(2,col_index);
      let columnName = curr_range.getA1Notation().replace(/[0-9]/g, '');
      activeSheet.getRange(columnName+":"+columnName).setBackground('#f27f49');
    }
    
    col_index++;
  }

  //generate month column
  col_index = default_sheet_col.findIndex(is_calendar)+1;
  for(let i=0; i<date_list.length; i++){
    activeSheet.getRange(1,col_index).setValue(monthNames[date_list[i].getMonth()]);
    col_index++;
  }

  Logger.log("Sheet Generated");
}

//utility functions

const is_calendar = (col_name) => col_name ==="{calendar}";

Date.prototype.addDays = function(days) {
    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}

function getDates(startDate, stopDate) {
    var dateArray = new Array();
    var currentDate = startDate;
    while (currentDate <= stopDate) {
        dateArray.push(new Date (currentDate));
        currentDate = currentDate.addDays(1);
    }
    return dateArray;
}

function colIndexToName(idx){
  let result =  String.fromCharCode(idx+65);
  Logger.log(result);
  return result;
}