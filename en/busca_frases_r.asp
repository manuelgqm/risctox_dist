<!--#include file="../dn_conexion.asp"-->
<%

	frase_r = EliminaInyeccionSQL(request("id"))
	texto = ""
	'response.write frase_r
	if frase_r="todas" then
		sql = "SELECT * FROM dn_risc_frases_r ORDER BY frase"
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		objRecordset.Open sql,objConnection,adOpenKeyset
		do while not objRecordset.eof
			texto=texto&"<tr><td class=celda3 align=right><b>"&objRecordset("frase")&"</b></td><td class=celda3 align=left>"&objRecordset("texto_ing")&"</td></tr>"
			objRecordset.movenext
		loop

	else
		frase_r = replace(frase_r,"Carc.Cat.1","")
		frase_r = replace(frase_r,"Carc.Cat.2","")
		frase_r = replace(frase_r,"Carc.Cat.3","")
		frase_r = replace(frase_r,"Carc.Cat. 1","")
		frase_r = replace(frase_r,"Carc.Cat. 2","")
		frase_r = replace(frase_r,"Carc.Cat. 3","")
		frase_r = replace(frase_r,"Repr.Cat.1","")
		frase_r = replace(frase_r,"Repr.Cat.2","")
		frase_r = replace(frase_r,"Repr.Cat.3","")
		frase_r = replace(frase_r,"Repr. Cat. 1","")
		frase_r = replace(frase_r,"Repr. Cat. 2","")
		frase_r = replace(frase_r,"Repr. Cat. 3","")
		frase_r = replace(frase_r," ","-")
		frase_r = replace(frase_r," ","-")
		frase_r = replace(frase_r," ","-")
		frase_r = replace(frase_r,"%20","-")
		frase_r = replace(frase_r,"R","")
		frase_r = replace(frase_r,";","-")
		frase_r = replace(frase_r,":","-")

		partes = split(frase_r,"-")

		for i=0 to Ubound(partes)
			sql = "SELECT * FROM dn_risc_frases_r WHERE frase='R"&partes(i)&"'"
			set objRecordset = Server.CreateObject ("ADODB.Recordset")
			objRecordset.Open sql,objConnection,adOpenKeyset
			if not objRecordset.eof then texto=texto&"<tr><td class=celda3 align=right><b>"&objRecordset("frase")&"</b></td><td class=celda3 align=left>"&objRecordset("texto_ing")&"</td></tr>"
		next
	end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>RISCTOX: R-Phrases</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="Risctox" />
<meta name="Author" content="SPL Sistemas de Informaci&oacute;n - www.spl-ssi.com" />
<meta name="description" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Subject" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Keywords" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Language" content="English" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />

<link rel="stylesheet" type="text/css" href="../estructura.css">
<link rel="stylesheet" type="text/css" href="../dn_estilos.css">
<link rel="stylesheet" type="text/css" href="css/en.css">

<body>
<table><tr><td class="titulo3">R-Phrases</td></tr>
<tr><td class="campo">Nature of special risks attributed to dangerous substances and preparations</td></tr>
</table>
&nbsp;
<table class="tabla3" width="90%" align="center" height="100%" valign="middle">
<%=texto%>
</table>
<p align="center"><input type="button" class="boton" value="See all R-Phrases" onClick="location.href='busca_frases_r.asp?id=todas'"></p>
</body>
</html>