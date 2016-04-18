<!--#include file="../EliminaInyeccionSQL.asp"-->
<%
' MENSAJES DE ERROR
// ##########################################################################
function flashMsgShow()

	'Si existe la variable de sessión "flashMsg", la muestra y después la borra
	if (not session("flashMsg")="") then
%>
		<fieldset id="flashmsg"><legend class="<%=lcase(session("flashType"))%>"><strong><%=session("flashType")%></strong></legend><%=session("flashMsg")%></fieldset>
<%
		session("flashType")=""
		session("flashMsg")="" 
	end if
end function

function flashMsgCreate(msg, tipo)

	'Crea mensaje de error
	session("flashType")=tipo	
	session("flashMsg")=msg

end function

function comprobarl(valor,max,nombre)
	if len(valor)>max then
		comprobarl="<br />-Se ha sobrepasado la longitud máxima (" &max& ") para el campo " &nombre& ": " &valor
	else
		comprobarl=""
	end if
end function
%>

<%
'************BÚSQUEDA (TILDES)*********
function quitartildes(byVal termino)
	if (isnull(termino)) then
		quitartildes = null
	else

		' Reemplazamos todas las tildes.
		termino = replace(termino,"á","a")
		termino = replace(termino,"é","e")
		termino = replace(termino,"í","i")
		termino = replace(termino,"ó","o")
		termino = replace(termino,"ú","u")	

		termino = replace(termino,"à","a")
		termino = replace(termino,"è","e")
		termino = replace(termino,"ì","i")
		termino = replace(termino,"ò","o")
		termino = replace(termino,"ù","u")
	
		termino = replace(termino,"ü","u")
	
		termino = replace(termino,"Á","A")
		termino = replace(termino,"É","E")
		termino = replace(termino,"Í","I")
		termino = replace(termino,"Ó","O")
		termino = replace(termino,"Ú","U")		

		termino = replace(termino,"À","A")
		termino = replace(termino,"È","E")
		termino = replace(termino,"Ì","I")
		termino = replace(termino,"Ò","O")
		termino = replace(termino,"Ù","U")
	
		termino = replace(termino,"Ü","U")

		quitartildes=termino
	end if
end function

function montartildes(byVal termino)

	' pasamos a exp. regular con todas las posibilidades

	termino = replace(termino,"a","[aáà]")
	termino = replace(termino,"e","[eéè]")
	termino = replace(termino,"i","[iíì]")
	termino = replace(termino,"o","[oóò]")
	termino = replace(termino,"u","[uúùü]")

	montartildes=termino

end function
%>

<%
' FORMULALIOS
' ##########################################################################
FUNCTION dimechecked (booleano)
	dimechecked=""
	if booleano<>"" then
		if booleano=true or booleano=1 then response.write " checked='checked' "
	end if
END FUNCTION
%>

<%
' ARCHIVOS
' ##########################################################################
SUB borrarfichero (archivo)

	archivo=server.mappath(".")&archivo
	response.write " - archivo a borrar: " &archivo
	Set fso = CreateObject("Scripting.FileSystemObject")
	if (fso.FileExists(archivo)) then fso.DeleteFile archivo
	set fso = nothing
	
end sub

Function RenombrarArchivo(origen,destino)
	
	destinosinruta=destino
	origen=server.mappath(".")&origen
	destino=server.mappath(".")&destino
	'response.write "<p>origen: " &origen
	'response.write "<p>destino: " &destino
	Set fso = CreateObject("Scripting.FileSystemObject")
	if (fso.FileExists(origen)) then
		'si existe ya destino, hay que eliminarlo o dara un error
		if (fso.FileExists(destino)) then borrarfichero (destinosinruta)
		fso.MoveFile origen, destino
	end if
	set fso = nothing
	
End function

function dimenuevonombre(nombre,extension)
	
	if extension="?" then
		dimenuevonombre="KO"
	else
		dimenuevonombre= nombre& "." &extension
	end if
	
end function

function dimeextension(nombrearchivo)

	'DEVUELVE EXTENSION, en minusculas, y sin punto
	'si dimeextension devuelve 0, es que no habia extension (no habia punto)
	'si devuelve 2, es que habia mas de un punto (el nombre de archivo no se considera válido)
	'si devuelve 3, es que habia extension, pero de menos de dos caracteres, por lo que no se considera valida
	'en otro caso, es que se ha podido determinar la extension, y sera lo que devuelva 

	'primero, comprobamos que hay un punto, ni mas ni menos
	nombretemp=nombrearchivo
	contador=0
	do while instr(nombretemp,".")
		contador=contador+1
		caracteresapartirdeestepunto=len(nombretemp)-instr(nombretemp,".")
		nombretemp=right(nombretemp,caracteresapartirdeestepunto)
	loop
	
	select case contador
		case 0: dimeextension=0
		case 1: punto=instr(nombrearchivo,".")
				exten=right(nombrearchivo,len(nombrearchivo)-punto)
				if len(exten)>=2 then
					dimeextension=lcase(exten)
				else
					dimeextension=3
				end if
		case else: dimeextension=2
	end select

end function
%>

<%
'VARIOS************************************

function roundsup (pes) 'Redondea a enteros superiores.

	ncoma=instr(pes, ",")
	if ncoma>0 then
	parteentera=int(pes)
	pes=parteentera+1
	end if
	roundsup=(pes)

end function

function quitaultimoscar(cadena,ncars)
	quitaultimoscar=cadena
	if cadena<>"" then
		if len(cadena)>ncars then	quitaultimoscar=left(cadena,(len(cadena)-ncars))
	end if
end function

' ##########################################################################

function dameSinonimos(byval id_sustancia)
	' Devuelve lista de sinónimos para la sustancia indicada
	sinonimos = ""

	sql="SELECT nombre FROM dn_risc_sinonimos WHERE id_sustancia="&id_sustancia&" ORDER BY nombre"
	set objRst=objConn1.execute(sql)
	if (not objRst.eof) then
		sinonimos = sinonimos & "<ul>"
		do while (not objRst.eof)
			sinonimos = sinonimos &"<li>"&corta(objRst("nombre"), 90, "puntossuspensivos")&"</li>"
			objRst.movenext
		loop
		sinonimos = sinonimos & "</ul>"
	end if
	objRst.close()
	set objRst=nothing

	dameSinonimos = sinonimos
end function

' ##########################################################################

sub log_this(byval descripcion)
  ' Apunta en la tabla dn_logs el mensaje indicado, junto con el script y la fecha
  script = Request.ServerVariables("SCRIPT_NAME")
  fecha = now
  ' Si la descripcion es muy larga, la trunca
  if (len(descripcion) > 1000) then
    descripcion = left(descripcion, 1000)
  end if
  sql = "INSERT INTO dn_log (script, fecha, descripcion) VALUES ('"&script&"','"&fecha&"','"&descripcion&"')"
  objConn1.execute(sql),,adexecutenorecords
end sub
%>

