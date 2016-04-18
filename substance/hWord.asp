<!--#include virtual="config/checkAccess.asp"-->
<!--#include virtual="config/dbConnection.asp"-->
<!--#include virtual="lib/requestHelper.asp"-->
<%

	frase_h = requestClean("id")
	texto = ""
	if frase_h="todas" then
		sql = "SELECT * FROM dn_risc_frases_h ORDER BY frase"
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		objRecordset.Open sql, connStr, adOpenKeyset
		do while not objRecordset.eof
			texto=texto&"<tr><td class=celda3 align=right><b>"&objRecordset("frase")&"</b></td><td class=celda3 align=left>"&objRecordset("texto")&"</td></tr>"
			objRecordset.movenext
		loop

	else
		
		frase_h = replace(frase_h,"-","")
		partes = split(frase_h,";")
	
		for i=0 to Ubound(partes)
			sql = "SELECT * FROM dn_risc_frases_h WHERE frase='" & trim(partes(i)) & "'"
			set objRecordset = Server.CreateObject ("ADODB.Recordset")
			objRecordset.Open sql, connStr, adOpenKeyset
			if not objRecordset.eof then texto=texto&"<tr><td class=celda3 align=right><b>"&objRecordset("frase")&"</b></td><td class=celda3 align=left>"&objRecordset("texto")&"</td></tr>"
		next
	end if
	
	objRecordset.Close
	set objRecordset = nothing
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Frases H</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="SPL Sistemas de Información - www.spl-ssi.com" />
<meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<link rel="stylesheet" type="text/css" href="../estructura.css"  />
<body>
<table><tr><td class="titulo3">Frases H</td></tr>
<tr><td class="campo">Naturaleza de los riesgos específicos atribuidos a las sustancias y preparados peligrosos</td></tr>
</table>
&nbsp;
<table class="tabla3" width="90%" align="center" height="100%" valign="middle">
<%=texto%>
</table>
<p align="center"><input type="button" class="boton" value="VER TODAS LAS FRASES H" onClick="location.href='hWord.asp?id=todas'"></p>
</body>
</html>