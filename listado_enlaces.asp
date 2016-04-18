<!--#include file="web_inicio.asp"-->

<html>
<head>
<title>Listado de enlaces</title>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
</head>

<body bgcolor="#F4F2B8" topmargin="10" leftmargin="15">


<p class=negro>Pulsa un enlace para pasarlo a la ventana de edición de página</p>
<table class="negroindice">
<%
	Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
	ordensql2 = "SELECT id,titulo FROM ENL_ENLACES ORDER BY id DESC"
	objrecordset2.Open ordensql2,OBJConnection,adOpenKeyset
	i = 1
	do while not objrecordset2.eof
		i = -1*i
		if i=1 then 
		  	color = "#FFFFB3"
	  	else
		  	color = "#FFFFD2"
		end if
%>
<tr><td bgcolor="<%=color%>" nowrap><a onclick="window.opener.formulario.enlaces.value='<%=objrecordset2("id")%>';window.close();" style="cursor:hand"><%=objrecordset2("id")&". "&objrecordset2("titulo")%></a></td></tr>
<%  	objrecordset2.movenext
	loop
%>
</table>

</body>
</html>
