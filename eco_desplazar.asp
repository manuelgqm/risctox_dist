<!--#include file="eco_conexion.asp"-->
<% 	

	id = EliminaInyeccionSQL(request("id"))
	sql = "SELECT numeracion FROM WEBISTAS_PAGINAS WHERE idpagina="&id
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	numeracion = objRecordSet("numeracion")
	nivelinicial = len(numeracion)
	for i=2 to 10
		num = EliminaInyeccionSQL(request("c"&i))
		'response.write "c"&i&"="&num&"<br>"
		if cstr(num)<>"" and i>2 then nuevanumeracion = nuevanumeracion & cstr(num) & "."
		if cstr(num)<>"" then letras = letras & chr(clng(num)+64)
	next
	'response.write letras&"<br>"&numeracion

	sql = "SELECT idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='"&letras&"'"
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,OBJConnection,adOpenKeyset
	repes = objrecordset.recordcount

	sql = "SELECT numeracion,titulo FROM WEBISTAS_PAGINAS WHERE numeracion LIKE '"&numeracion&"%' ORDER BY numeracion"
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,OBJConnection,adOpenKeyset

%>
<HTML>

<head>
<title>Desplazar la página <%=id%></title>
<base target="_self">

</head>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
<body bgcolor="#FEFEFE" topmargin="20" leftmargin="20">
<script LANGUAGE="JScript">
<!--
// 
function mover()
{
	location.href = 'eco_desplazar2.asp?id=<%=id%>&letras=<%=letras%>';
}
//-->
</script>
<% if repes<>0 then %>
<p class="cue_fuente">ATENCIÓN: existe otra página con la misma clasificación</p>
<p align="center"><input type="button" value="VOLVER" onClick="location.href='eco_editarpagina.asp?id=<%=id%>'" class="boton"></p>
<% else %>
<p class="cue_fuente">Desplazando las páginas:</p>
<table>
<% do while not objrecordset.eof %>
<tr>
<td class="cue_fuente">
<% 
   numeracion = objrecordset("numeracion")
   nivel = len(numeracion)
   espacio = ""
   situacion = asc(mid(numeracion,3,1))-64
   for i = 3 to nivel
    espacio = espacio&"&nbsp;&nbsp;&nbsp;&nbsp;"
    if i<>3 then situacion = situacion&"."&(asc(mid(numeracion,i,1))-64)
   next
   response.write espacio & situacion & ". " & objrecordset("titulo")
   objrecordset.movenext
%>
</td></tr>
<% loop %>
</table>
<p class="cue_fuente">...a la numeración:</p>
<table>
<% objrecordset.movefirst
   do while not objrecordset.eof %>
<tr>
<td class="cue_fuente">
<% numeracion = objrecordset("numeracion")
   nivel = len(numeracion)
   espacio = ""
   situacion = ""
   for i = nivelinicial+1 to nivel
    espacio = espacio&"&nbsp;&nbsp;&nbsp;&nbsp;"
    situacion = situacion & (asc(mid(numeracion,i,1))-64) & "."
   next
   response.write espacio & nuevanumeracion & situacion & " "& objrecordset("titulo")
   objrecordset.movenext
%>
</td></tr>
<% loop %>
</table>
<p align="center"><input type="button" value="ACEPTAR" onClick="mover()" class="boton">&nbsp;&nbsp;
<input type="button" value="CANCELAR" onClick="location.href='eco_editarpagina.asp?id=<%=id%>'" class="boton"></p>
<% end if %>
</body>
</HTML>