function valida_mail(campo){
    
// validar la cuenta de correo usando una expresi�n regular (RegExp)
if(campo.value!='')
 if(campo.value.search(/^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$/ig)){
   alert("La cuenta de correo introducida no es v�lida, debes escribirla de forma: nombre@servidor.dominio");
   campo.select();
   campo.focus();
   return false;
  }
}
