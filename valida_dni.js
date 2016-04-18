function valida_dni(campo)
{
 if(campo.value!='')
 {
 numeros_dni = "";
 numeros_dni = campo.value;
 if ((numeros_dni.charAt(0)!="X") && (numeros_dni.charAt(0)!="x") && (numeros_dni.charAt(0)!="P") && (numeros_dni.charAt(0)!="p") && isNaN(numeros_dni.charAt(0)))
 	{  alert("El primer número debe estar entre 0 y 9 o ser una letra X o P");
	   campo.focus(); 
	   return (false);
	   }
 else
 	if (isNaN(numeros_dni.substring(1,numeros_dni.length)))
 	{  alert("Tienes que escribir los números de tu DNI/NIE sin puntos ni símbolos");
 	   campo.focus();
	   return (false); 	   
 	    }
 	else
 	
		if (numeros_dni.length!=8) 
		{  alert("Hay que escribir 8 caracteres en el DNI/NIE");
	   	   campo.focus();
	           return (false);	   	   
	   	    }
		
 	
}

num_calcula = numeros_dni;

if ((numeros_dni.charAt(0)=="X") || (numeros_dni.charAt(0)=="x") || (numeros_dni.charAt(0)=="P") || (numeros_dni.charAt(0)=="p"))
 num_calcula = num_calcula.substring(1,8);
 
cadena="TRWAGMYFPDXBNJZSQVHLCKET"
posicion = num_calcula % 23
letra = cadena.substring(posicion,posicion+1)

campo.value = numeros_dni + letra;

}




