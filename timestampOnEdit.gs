// Used this code to create an edit timestamp in the cell next to the edited cell. Used in combo with highlighting and text colour changes to hide the timestamp when the cell is empty. 

// NB: Update column and row numbers as required in lines 11 and 13. Lines 11-15 can be repeated for multiple timestamps per sheet, with the column and row numbers updated accordingly.

function timestampOnEdit(e) {

  var row = e.range.getRow();

  var col = e.range.getColumn();

  if(col == 2, row == 42){

    e.source.getActiveSheet().getRange(42,3).setValue(new Date());

  }
