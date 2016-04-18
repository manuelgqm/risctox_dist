function valida_numero(fieldName) {

fieldValue = fieldName.value;
if (isNaN(fieldValue)) {
  alert("Introduce un número!!!.");
  fieldName.select();
  fieldName.focus();
  return (false);
}

return (true);  

}
