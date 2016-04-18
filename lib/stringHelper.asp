<%
' ################################################
function h(byval cadena)
	' Reemplaza caracteres especiales por su correspondiente c�digo HTML
	if (isNull(cadena)) then
		h = ""
	else
		' Quitamos espacios al principio y final
		cadena = trim(cadena)

		' Evitar SQL Injection
		cadena = Replace(cadena,"�","'") 	' Ap�strofe
		cadena = Replace(cadena,"'","&#39;") 	' Ap�strofe
		cadena = Replace(cadena,"�","&#39;") 	' Ap�strofe
		cadena = Replace(cadena,"""","&#34;") ' La comilla doble (debe ir doble en asp)
		cadena = Replace(cadena,"%","&#37;") 	' Porcentaje
		cadena = Replace(cadena,"[","&#91;")	' Corchete izq
		cadena = Replace(cadena,"]","&#93;")	' Corchete dch

		' Mayor y menor
		cadena = Replace(cadena,"<", "&lt;")
		cadena = Replace(cadena,">", "&gt;")

		' Tildes
		'cadena = Replace(cadena,"�","&aacute;")
		'cadena = Replace(cadena,"�","&eacute;")
		'cadena = Replace(cadena,"�","&iacute;")
		'cadena = Replace(cadena,"�","&oacute;")
		'cadena = Replace(cadena,"�","&uacute;")

		'cadena = Replace(cadena,"�","&Aacute;")
		'cadena = Replace(cadena,"�","&Eacute;")
		'cadena = Replace(cadena,"�","&Iacute;")
		'cadena = Replace(cadena,"�","&Oacute;")
		'cadena = Replace(cadena,"�","&Uacute;")

		'cadena = Replace(cadena,"�","&agrave;")
		'cadena = Replace(cadena,"�","&egrave;")
		'cadena = Replace(cadena,"�","&igrave;")
		'cadena = Replace(cadena,"�","&ograve;")
		'cadena = Replace(cadena,"�","&ugrave;")

		'cadena = Replace(cadena,"�","&Agrave;")
		'cadena = Replace(cadena,"�","&Egrave;")
		'cadena = Replace(cadena,"�","&Igrave;")
		'cadena = Replace(cadena,"�","&Ograve;")
		'cadena = Replace(cadena,"�","&Ugrave;")

		'cadena = Replace(cadena,"�","&ntilde;")
		'cadena = Replace(cadena,"�","&Ntilde;")

		'cadena = Replace(cadena,"�","&uuml;")
		'cadena = Replace(cadena,"�","&Uuml;")

		' Y ya. Devolvemos la cadena corregida
		h = cadena
	end if
end function

' ###############################################

function hjs(byval cadena)
	' Versi�n de h() para emplear en JavaScript. Sustituye ' por \'
	if (isNull(cadena)) then
		hjs = ""
	else
		' Quitamos espacios al principio y final
		cadena = trim(cadena)

		' Evitar SQL Injection
		cadena = Replace(cadena,"'","\'") 	' Ap�strofe
		cadena = Replace(cadena,"�","\'") 	' Ap�strofe
		'cadena = Replace(cadena,"�","\'") 	' Ap�strofe
		cadena = Replace(cadena,"&#39;","\'") 	' Ap�strofe

		' Y ya. Devolvemos la cadena corregida
		hjs = cadena
	end if
end function

' ###############################################

function h2(byval cadena)
	' Reemplaza caracteres codificados en HTML por el caracter de verdad, para deshacer
	' lo cambiado por la funci�n h
	if (isNull(cadena)) then
		h2 = ""
	else
		cadena = Replace(cadena,"&#34;","'") 	' Ap�strofe
		cadena = Replace(cadena,"&#39;","""") ' La comilla doble (debe ir doble en asp)
		cadena = Replace(cadena,"&#37;","%") 	' Porcentaje
		cadena = Replace(cadena,"&#91;","[")	' Corchete izq
		cadena = Replace(cadena,"&#92;","]")	' Corchete dch
		cadena = Replace(cadena,"&lt;", "<")
		cadena = Replace(cadena,"&gt;", ">")

		h2 = cadena
	end if
end function

' ###############################################

function quitartildes(byVal termino)
	' Reemplaza tildes por letra correspondiente sin tilde
	if (isnull(termino)) then
		quitartildes = null
	else

		' Reemplazamos todas las tildes.
		termino = replace(termino,"�","a")
		termino = replace(termino,"�","e")
		termino = replace(termino,"�","i")
		termino = replace(termino,"�","o")
		termino = replace(termino,"�","u")

		termino = replace(termino,"�","a")
		termino = replace(termino,"�","e")
		termino = replace(termino,"�","i")
		termino = replace(termino,"�","o")
		termino = replace(termino,"�","u")

		termino = replace(termino,"�","u")

		termino = replace(termino,"�","A")
		termino = replace(termino,"�","E")
		termino = replace(termino,"�","I")
		termino = replace(termino,"�","O")
		termino = replace(termino,"�","U")

		termino = replace(termino,"�","A")
		termino = replace(termino,"�","E")
		termino = replace(termino,"�","I")
		termino = replace(termino,"�","O")
		termino = replace(termino,"�","U")

		termino = replace(termino,"�","U")

		quitartildes=termino
	end if
end function

' ###############################################

function acortarCadena(cadena, max, sustitucion)
	' Si la longitud de la cadena es mayor que max,
	' nos quedamos con el inicio y el final y lo concatenamos con la sustitucion en medio
	nuevaCadena = cadena

	if (len(cadena) > max) then
		' Acortar
		' Calculamos la mitad
		longitudIzquierda = int(max/2)
		longitudDerecha = longitudIzquierda
		' Si max es impar, entonces le damos 1 m�s a la parte izquierda
		if ((max mod 2) = 1) then
			longitudIzquierda = longitudIzquierda + 1
		end if

		parteIzquierda = left(cadena,longitudIzquierda)
		parteDerecha = right(cadena,longitudDerecha)

		nuevaCadena = parteIzquierda&sustitucion&parteDerecha
	end if

	' Devolvemos la cadena
	acortarCadena = nuevaCadena
end function

' ################################################

function corta (cadena, maximo, quehacer)	'corta una cadena larga, y le podemos decir longitud y lo que tiene que hacer: puntossuspensivos, meterbrs
	lencadena=len(cadena)	if lencadena>maximo then		select case quehacer			case "puntossuspensivos": 	cadena=left(cadena,maximo-3) & "..."			case "meterbrs": 			nbrs=lencadena/maximo										cadena=metebr (cadena, maximo)										'for i=0 to cint(nbrs)											'response.write "<br>i " &i											'reemplazado=reemplazado&Replace(cadena," ","<br />",maximo*(i+1),1)											'reemplazado=metebr(porreemplazar,maximo)											'porreemplazar=right(cadena,lencadena-(maximo*(i+1)))										'next		end select	end if	corta=cadenaend function


function metebr (micadena, mimaximo)	'auxiliar de corta: en una cadena, mete un br en maximo (la anterior la llamara tantas veces como sea preciso)	'se podria mejorar haciendo que solo parta cuando encuentre un espacio	milencadena=len(micadena)	'response.write "<br>milencadena " &milencadena	'response.write "<br>mimaximo" &mimaximo	if milencadena>mimaximo then		izq=left(micadena,mimaximo)		der=replace(micadena,izq,"")		micadena=metebr(izq, mimaximo) &"<br />"& metebr(der, mimaximo)	end if	metebr=micadenaend function

' ################################################
function comprobarReferer(byval refererEsperado)
	' Comprueba que a la p�gina de proceso
	' se llega desde un formulario con el nombre indicado, en el mismo dominio y ruta

	nombreProceso=request.servervariables("SCRIPT_NAME") ' /clientes/prueba.asp
	posicionBarra=instrrev(nombreProceso,"/") ' desde la derecha
	nombreDirectorio=left(nombreProceso,posicionBarra) ' /clientes/
	nombreDominio = request.servervariables("SERVER_NAME")

	urlRefererReal = request.servervariables("HTTP_REFERER")
	' Al referer real le quitamos los par�metros por URL, a partir de "?"
	posInterrogacion=instr(urlRefererReal,"?")
	if (posInterrogacion > 0) then
		urlRefererReal=left(urlRefererReal,posInterrogacion-1)
	end if
	urlRefererEsperado = "http://"&nombreDominio&nombreDirectorio&refererEsperado

	if (urlRefererEsperado = urlRefererReal) then
		' Coincide, devolvemos true
		comprobarReferer=true
	else
		' No coincide, devolvemos false
		comprobarReferer=false
	end if
end function

' ###############################################
function comprobarLen(nombreCampo,campo,maxLen,obligatorio)
	' Para comprobar que no se sobrepase la longitud de cada campo
	' nombrecampo -- nombre del campo, para el usuario
	' campo -- el campo a comprobar
	' maxLen -- longitud m�xima
	' obligatorio -- si es obligatorio o no

	menError = "" ' Inicialmente no hay mensaje de error
	if  campo<>"" then
		if len(campo)>maxLen then menError="<br/>- El campo " &nombreCampo& " no debe pasar de los " &maxLen& " caracteres."
	else 'si esta vacio y es obligatorio, tb es un error
		if obligatorio then menError="<br/>- El campo " &nombreCampo& " es obligatorio."
	end if

	comprobarLen=menError
end function

' ###############################################
function comprobarLong(cadena, min, max, nombreCampo, obligatorio)
	' Comprueba que la cadena pasada es un n�mero entero entre min y max
	' El tipo de datos Long va desde (-2^31 - 1) hasta (2^31 - 1), como el tipo int de SQL Server
	' Para poder comprobar que estamos dentro del rango es preciso convertir a double
	menError = "" ' Inicialmente no hay mensaje de error
	if (cadena="") then
		if (obligatorio) then
			menError="\n- El campo " &nombreCampo& " es obligatorio."
		end if
	else
		if (isnumeric(cadena) and (cadena<>"")) then
			if ((cdbl(cadena) >= -2^31-1) and (cdbl(cadena) <= 2^31-1)) then
				if ((clng(cadena) < min) or (clng(cadena) > max)) then
					menError = "\n- El campo "&nombreCampo&" debe estar entre "&min&" y "&max&"."
				end if
			else
				menError = "\n- El campo "&nombreCampo&" est� fuera de los rangos permitidos."
			end if
		else
			menError = "\n- El campo "&nombreCampo&" debe ser un n�mero entero entre "&min&" y "&max&"."
		end if
	end if

	comprobarLong=menError
end function

' ##############################################
sub mailtoSinSpam(email,asunto)
	' Envio de correo: lo hacemos mediante funci�n JavaScript para evitar spambots
	' Tiene que estar incluido el JavaScript: enviarCorreo(usuario,dominio,extension,asunto)
	' Uso: incluir como ASP en evento onclick de enlaces, de esta guisa:
	' mailtoSinSpam responsable_email, "Desde www.publicaciones-isp.org"

	' Usuario (hasta @)
	posicionArrobaIzq=instr(email,"@")
	usuario=left(email,posicionArrobaIzq-1)

	' Extensi�n (desde �ltimo punto)
	longitudEmail=len(email)
	posicionPuntoDch=instrrev(email,".")
	extension=right(email,longitudEmail-posicionPuntoDch)

	' Dominio (desde @ hasta punto)
	longitudUsuario=len(usuario)
	longitudExtension=len(extension)
	dominioConExtension=right(email,longitudEmail-(longitudUsuario+1))
	dominio=left(dominioConExtension,len(dominioConExtension)-(longitudExtension+1))

	' Cadena mailto final
	mailto="enviarCorreo('"&usuario&"', '"&dominio&"', '"&extension&"', '"&asunto&"')"

	' Devolvemos la cadena
	response.write mailto
end sub


' ###################################################
function claveAleatoria(longitud)
	' Empezamos el random...
	randomize
	' cambiar el valor de la siguiente variable para...
	' modificar la longitud del c�digo que generaremos
	for contador = 1 to longitud
		' hacemos random entre 97 y 122.
		numero = Int(26 * Rnd + 97)

		' tomamos el numero y lo cambiamos por la letra
		letra = Chr(numero)

		' Agregamos la nueva letra
		codigo = codigo & letra
	next

	' devolvemos el c�digo
	claveAleatoria=codigo
end function

' #########################################
function soloAlfanumerico(cadena)
	' Devuelve la cadena limpia para evitar caracteres raros y proteger frente a
	' SQL Injection

	'Create a regular expression object
	Dim regEx
	Set regEx = New RegExp
	'The global property tells the RegExp engine to find ALL matching
	'substrings, instead of just the first instance. We need this to be true.
	regEx.Global = true
	'Our pattern tells us what to find in the string... In this case, we find
	'anything that isn't a numerical character, or a lowercase or
	'uppercase alphabetic character
	regEx.Pattern = "[^0-9a-z�A-Z�]"
	'Use the replace function of RegExp to clean the username. The replace
	'function takes the string to search (using the Pattern above as the
	'search criteria), and the string to replace any found strings with.
	'In this case, we want to replace our matches with nothing (''),
	'as the matching characters will be the ones we don't want in our username.
	soloAlfanumerico = LCase(regEx.Replace(cadena, ""))
end function

' #########################################
function soloNumerico(cadena)
	' Deja pasar s�lo digitos

	'Create a regular expression object
	Dim regEx, cadenaFinal
	Set regEx = New RegExp
	'The global property tells the RegExp engine to find ALL matching
	'substrings, instead of just the first instance. We need this to be true.
	regEx.Global = true
	'Our pattern tells us what to find in the string... In this case, we find
	'anything that isn't a numerical character, or a lowercase or
	'uppercase alphabetic character
	regEx.Pattern = "[^0-9]"
	'Use the replace function of RegExp to clean the username. The replace
	'function takes the string to search (using the Pattern above as the
	'search criteria), and the string to replace any found strings with.
	'In this case, we want to replace our matches with nothing (''),
	'as the matching characters will be the ones we don't want in our username.

	' Si no han quedado numeros tras la limpieza, devuelvo "0"
	cadenaFinal=LCase(regEx.Replace(cadena, ""))
	if (cadenaFinal="") then
		cadenaFinal="0"
	end if

	soloNumerico = clng(cadenaFinal)
end function

' ###################################################

function es_frase(byval cadena,byval tipo)
	' Devuelve booleano si la cadena pasada tiene pinta de frase R, o sea, primer caracter es R y segundo, de 1 a 9
	if (len(cadena) >= 2) then
		' Longitud 2 o m�s
		caracter_1 = mid(cadena,1,1)
		caracter_2 = mid(cadena,2,1)

		if ((caracter_1 = tipo) and (instr("123456789",caracter_2)>0)) then
			es_frase=true
		else
			es_frase=false
		end if
	else
		es_frase=false
	end if
end function

' ###################################################

function espaciar (byval cadena)
	' Mete espacios en nombres para que se puedan partir sin necesidad de br
	cadena = replace (cadena, "-", " - ")
	cadena = replace (cadena, "(", " (")
	cadena = replace (cadena, ")", ") ")
	cadena = replace (cadena, "[", " [")
	cadena = replace (cadena, "]", "] ")
	cadena = replace (cadena, ",", ", ")
	cadena = replace (cadena, ".", ". ")
	cadena = replace (cadena, "  ", " ")
	cadena = replace (cadena, "  ", " ")
	cadena = replace (cadena, "  ", " ")
	cadena = replace (cadena, "  ", " ")
  cadena = replace (cadena, "@", "@ ")

	espaciar = cadena
end function

Function nl2br(text)
  nl2br = replace(text, vbNewLine, "<br />")
End Function

' ###################################################

function elimina_repes(byval cadena, byval separador)
  nueva=""

  vector=split(cadena, separador)
  for i=0 to ubound(vector)
    if (nueva = "") then
      nueva = vector(i)
    else
      ' Solo se a�ade si no existe ya
      if (instr(lcase(nueva), lcase(vector(i))) = 0) then
        nueva = nueva & separador & vector(i)
      end if
    end if
  next

  elimina_repes = nueva
end function



' ####################################################33
function ereg_replace(strOriginalString, strPattern, strReplacement, varIgnoreCase)
	' Function replaces pattern with replacement
	' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
	dim objRegExp : set objRegExp = new RegExp
	with objRegExp
		.Pattern = strPattern
		.IgnoreCase = varIgnoreCase
		.Global = True
	end with
	ereg_replace = objRegExp.replace(strOriginalString, strReplacement)
	set objRegExp = nothing
end function










%>
