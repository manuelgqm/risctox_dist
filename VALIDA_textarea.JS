function checkMaxLength (textarea, evt, maxLength) {
  if (textarea.selected && evt.shiftKey) 
    // ignore shift click for select
    return true;
  var allowKey = false;
  if (textarea.selected && textarea.selectedLength > 0)
    allowKey = true;
  else {
    var keyCode = 
      document.layers ? evt.which : evt.keyCode;
    if (keyCode < 32 && keyCode != 13)
      allowKey = true;
    else           
      allowKey = textarea.value.length < maxLength;
  }
  textarea.selected = false;
  return allowKey;
}
function storeSelection (field) {
  if (document.all) {
    field.selected = true;
    field.selectedLength = 
      field.createTextRange ?
        document.selection.createRange().text.length : 1;
  }
}