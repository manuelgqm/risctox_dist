<!--#include file="eco_conexion.asp"-->
<html>
<head>
<title>Listado de páginas</title>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
</head>

<body bgcolor="#FEFEFE" topmargin="10" leftmargin="15">


<p class=negro>Pulsa una página para pasarla a la ventana de edición de página</p>
<table class="negroindice">
<%
	Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
	ordensql2 = "SELECT idpagina,titulo,numeracion FROM WEBISTAS_PAGINAS WHERE numeracion LIKE 'AI%' ORDER BY numeracion"
	objrecordset2.Open ordensql2,OBJConnection,adOpenKeyset
	i = 1
	do while not objrecordset2.eof
		titulo = ""
		numeracion = objrecordset2("numeracion")
		for i = 3 to len(numeracion)
			titulo = titulo & cstr(asc(mid(numeracion,i,1))-64) & "."
		next
		titulo = titulo & " " & objrecordset2("titulo")
		i = -1*i
		if i=1 then 
		  	color = "#FFFFB3"
	  	else
		  	color = "#FFFFD2"
		end if
%>
<tr><td bgcolor="<%=color%>" nowrap><a onclick="window.opener.formulario.listadopaginas.value='<%=objrecordset2("idpagina")%>';window.close();" style="cursor:hand"><%=titulo%></a></td></tr>
<%  	objrecordset2.movenext
	loop
%>
</table>

</body>
</html>
