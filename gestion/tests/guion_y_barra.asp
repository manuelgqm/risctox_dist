<!--#include file="../dn_conexion.asp"-->
<!--#include file="../dn_fun_texto.asp"-->
<!--#include file="../dn_fun_comunes.asp"-->
<!--#include file="../importador/fun_importador.asp"-->

<%
on error resume next
dim sql
response.buffer=true
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-15">
<style type="text/css">
body { font: x-small Verdana, Arial, Helvetica, sans-serif; }
</style>
</head>
<body>
<h1>Listado de las sustancias que tienen guión y barra en frases R en campos clasificacion_x</h1>

<table border="1">
	<tr>
		<th>Sustancia</th>
		<th>Clasificación 1</th>
		<th>Clasificación 2</th>
		<th>Clasificación 3</th>
		<th>Clasificación 4</th>
		<th>Clasificación 5</th>
		<th>Clasificación 6</th>
	</tr>
<%
response.buffer=true

sql = "SELECT * FROM dn_risc_sustancias WHERE (clasificacion_1 LIKE '%/%' AND clasificacion_1 LIKE '%-%') OR (clasificacion_2 LIKE '%/%' AND clasificacion_2 LIKE '%-%') OR (clasificacion_3 LIKE '%/%' AND clasificacion_3 LIKE '%-%') OR (clasificacion_4 LIKE '%/%' AND clasificacion_4 LIKE '%-%') OR (clasificacion_5 LIKE '%/%' AND clasificacion_5 LIKE '%-%') OR (clasificacion_6 LIKE '%/%' AND clasificacion_6 LIKE '%-%')"
'response.write "<tr><td colspan='8'>"&sql&"</td></tr>"

set objRst = objConn1.execute(sql)
if (objRst.eof) then
	response.write "<tr><td>No se ha encontrado ninguna sustancia con esas características</td></tr>"
else
	do while (not objRst.eof)
%>

	<tr>
		<td><a href="http://www.istas.net/ecoinformas/web/pruebas/web/dn_risctox_ficha_sustancia.asp?id_sustancia=<%= objRst("id") %>" target="_blank"><%= objRst("id") %></a></td>
		<td><%= resalta(objRst("clasificacion_1")) %></td>
		<td><%= resalta(objRst("clasificacion_2")) %></td>
		<td><%= resalta(objRst("clasificacion_3")) %></td>
		<td><%= resalta(objRst("clasificacion_4")) %></td>
		<td><%= resalta(objRst("clasificacion_5")) %></td>
		<td><%= resalta(objRst("clasificacion_6")) %></td>
	</tr>

<%
		objRst.movenext
		response.flush()
	loop
end if


' DESCONECTAMOS DEL SQL SERVER
cerrarconexion

' FIN
%>
</table>

<h1>FIN</h1>

</body>
</html>

<%
function resalta(byval cadena)
	' Si la cadena contiene "/" y "-", la muestra en rojo, si no, en gris
	if ((instr(cadena, "-") <> 0) and (instr(cadena, "/") <> 0)) then
		color = "red"
	else
		color = "grey"
	end if	
	resalta = "<font color='"&color&"'>"&cadena&"</font>"
end function
%>
