<!--#include file="EliminaInyeccionSQL.asp"-->
<!--#include file="dn_conexion.asp"-->

<%
 	Const adOpenKeyset = 1
'	DIM objConnection	
'	DIM objRecordset
	
'	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=disoltec02; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
'	OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
	
	OBJConnection.Open
	id = request("id")
	id = EliminaInyeccionSQL(id)
	texto = ""
	
	sql = "SELECT * FROM RQ_DEFINICIONES WHERE id="&id
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,objConnection,adOpenKeyset
	if not objRecordset.eof then texto=texto&"<tr><td class=titulo3 align=right valign=top>"&objRecordset("palabra")&"</td><td class=texto align=left>"&replace(objRecordset("definicion"),chr(13),"<br>")&"</td></tr>"
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Definición</title>
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
<body>
&nbsp;
<table class="tabla3" width="90%" align="center" height="100%" valign="middle" cellpadding="5">
<%=texto%>
</table>
&nbsp;
</body>
</html>