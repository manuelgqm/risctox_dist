<%
function es_frase_r(byval cadena)
	' Devuelve booleano si la cadena pasada tiene pinta de frase R, o sea, primer caracter es R y segundo, de 1 a 9
	if (len(cadena) >= 2) then
		' Longitud 2 o más
		caracter_1 = mid(cadena,1,1)
		caracter_2 = mid(cadena,2,1)

		if ((caracter_1 = "R") and (instr("123456789",caracter_2)>0)) then
			es_frase_r=true
		else
			es_frase_r=false
		end if
	else
		es_frase_r=false
	end if
end function

response.write es_frase_r("R0")
response.write es_frase_r("R01")
response.write es_frase_r("R10")
response.write es_frase_r("")
response.write es_frase_r("mi abuela")
%>
