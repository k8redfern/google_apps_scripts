// Used this code to create an edit timestamp in the cell next to the edited cell. Used in combo with highlighting and text colour changes to hide the timestamp when the cell is empty. 

function onEdit(e) {

  var row = e.range.getRow();

  var col = e.range.getColumn();

  if(col == 2, row == 42){

    e.source.getActiveSheet().getRange(42,3).setValue(new Date());

  }
