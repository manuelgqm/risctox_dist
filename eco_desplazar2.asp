<!--#include file="eco_conexion.asp"-->
<% 	

	id = EliminaInyeccionSQL(request("id"))
	letras = EliminaInyeccionSQL(request("letras"))
	sql = "SELECT numeracion FROM WEBISTAS_PAGINAS WHERE idpagina="&id
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	numeracion = objRecordSet("numeracion")
	nivelinicial = len(numeracion)

	sql2= "UPDATE WEBISTAS_PAGINAS SET numeracion='A"&letras&"'+right(numeracion,len(numeracion)-"&nivelinicial&") WHERE numeracion LIKE '"&numeracion&"%'"
	set objRecordset2 = OBJConnection.Execute(sql2)
	'response.write sql2
%>
<HTML>

<head>
<title>Desplazar la página <%=id%></title>
<base target="_self">

</head>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
<body bgcolor="#FEFEFE" topmargin="20" leftmargin="20">
<p class="cue_fuente">Desplazamiento realizado...</p>
<script LANGUAGE="JScript">
<!--
//
location.href='eco_editarpagina.asp?id=<%=id%>';
parent.frames.izquierda.location.reload();
//-->
</script>
</body>
</html>
