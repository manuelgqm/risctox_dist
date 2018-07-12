<!--#include file="../EliminaInyeccionSQL.asp"-->
<!--#include file="../dn_conexion.asp"-->

<%
 	Const adOpenKeyset = 1
'	DIM objConnection
'	DIM objRecordset

'	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=disoltec02; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED**"
'	OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED**"

	OBJConnection.Open
	id = request("id")
	id = EliminaInyeccionSQL(id)
	texto = ""

	sql = "SELECT * FROM RQ_DEFINICIONES WHERE id="&id
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,objConnection,adOpenKeyset
	if not objRecordset.eof then
		texto=texto&"<tr><td class=titulo3 align=right valign=top>"&objRecordset("palabra_eng")&"</td><td class=texto align=left>"&replace(objRecordset("definicion_eng"),chr(13),"<br>")&"</td></tr>"
	end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>RISCTOX: Definition</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="Risctox" />
<meta name="Author" content="SPL Sistemas de Información - www.spl-ssi.com" />
<meta name="description" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Subject" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Keywords" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Language" content="English" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />

<link rel="stylesheet" type="text/css" href="../estructura.css">
<link rel="stylesheet" type="text/css" href="css/en.css">

<body>
&nbsp;
<table class="tabla3" width="90%" align="center" height="100%" valign="middle" cellpadding="5">
<%=texto%>
</table>
&nbsp;
</body>
</html>