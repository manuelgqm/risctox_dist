function valida_numero(fieldName) {

fieldValue = fieldName.value;
if (isNaN(fieldValue)) {
  alert("Introduce un n�mero!!!.");
  fieldName.select();
  fieldName.focus();
  return (false);
}

return (true);  

}
