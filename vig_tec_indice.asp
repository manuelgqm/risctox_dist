<!--#include file="EliminaInyeccionSQL.asp"-->
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: <%=titulo%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multimèdia" />
<meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<link rel="stylesheet" type="text/css" href="estructura.css"  />
</head>
<body>
<% 
	sql = "SELECT id,afiliacion,titulo,url,indice FROM ENL_ENLACES WHERE id="&EliminaInyeccionSQL(request("id"))
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,objConnection,adOpenKeyset
%>
<table align="center" cellpadding="3" cellspacing="3">
<tr><td class="texto" valign="top" align="left"><a href="abreenlacer.asp?idenlace=<%=objRecordset("Id")%>" target="_blank"><img src="imagenes/ico_puntito.gif" valign="top" border=0></a></td><td class="texto" bgcolor="#DDDDDD"><b><%=objRecordset("afiliacion")%></b><br><a href="abreenlacer.asp?idenlace=<%=objRecordset("Id")%>" target="_blank"><%=objRecordset("titulo")%><br><a href="abreenlacer.asp?idenlace=<%=objRecordset("Id")%>" target="_blank"><%=objRecordset("url")%></a><br></td></tr>
<tr><td class="texto" valign="top" align="left" colspan="2"><%=replace(objRecordset("indice"),chr(13),"<br>")%></td></tr>
</table>
<p align="center"><input type="button" class="boton" value="CERRAR" onclick="window.close()"></p>

</body>
</html>