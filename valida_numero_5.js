function valida_numero_5(fieldName,maxvalor) {

fieldValue = fieldName.value;
if (isNaN(fieldValue) || (fieldValue>maxvalor)) {
  alert("Introduce un número en el margen 1-"+maxvalor+"!!!.");
  fieldName.select();
  fieldName.focus();
  return (false);
}

return (true);  

}
