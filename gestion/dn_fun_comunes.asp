<!--#include file="../EliminaInyeccionSQL.asp"-->
<%
' MENSAJES DE ERROR
// ##########################################################################
function flashMsgShow()

	'Si existe la variable de sessi�n "flashMsg", la muestra y despu�s la borra
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
		comprobarl="<br />-Se ha sobrepasado la longitud m�xima (" &max& ") para el campo " &nombre& ": " &valor
	else
		comprobarl=""
	end if
end function
%>

<%
'************B�SQUEDA (TILDES)*********
function quitartildes(byVal termino)
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

function montartildes(byVal termino)

	' pasamos a exp. regular con todas las posibilidades

	termino = replace(termino,"a","[a��]")
	termino = replace(termino,"e","[e��]")
	termino = replace(termino,"i","[i��]")
	termino = replace(termino,"o","[o��]")
	termino = replace(termino,"u","[u���]")

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
	'si devuelve 2, es que habia mas de un punto (el nombre de archivo no se considera v�lido)
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
	' Devuelve lista de sin�nimos para la sustancia indicada
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

