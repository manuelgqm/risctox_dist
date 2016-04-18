function valida_mail(campo){
    
// validar la cuenta de correo usando una expresión regular (RegExp)
if(campo.value!='')
 if(campo.value.search(/^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$/ig)){
   alert("La cuenta de correo introducida no es válida, debes escribirla de forma: nombre@servidor.dominio");
   campo.select();
   campo.focus();
   return false;
  }
}
