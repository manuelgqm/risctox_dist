<!--#include file="dn_conexion.asp"-->
<%

	frase_r = EliminaInyeccionSQL(request("id"))
	texto = ""
	'response.write frase_r
	if frase_r="todas" then
		sql = "SELECT * FROM RQ_FRASES_R ORDER BY frase"
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		objRecordset.Open sql,objConnection,adOpenKeyset
		do while not objRecordset.eof
			texto=texto&"<tr><td class=celda3 align=right><b>"&objRecordset("frase")&"</b></td><td class=celda3 align=left>"&objRecordset("texto")&"</td></tr>"
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
			sql = "SELECT * FROM RQ_FRASES_R WHERE frase='R"&partes(i)&"'"
			set objRecordset = Server.CreateObject ("ADODB.Recordset")
			objRecordset.Open sql,objConnection,adOpenKeyset
			if not objRecordset.eof then texto=texto&"<tr><td class=celda3 align=right><b>"&objRecordset("frase")&"</b></td><td class=celda3 align=left>"&objRecordset("texto")&"</td></tr>"
		next
	end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Frases R</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multim�dia" />
<meta name="description" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<link rel="stylesheet" type="text/css" href="estructura.css"  />
<body>
<table><tr><td class="titulo3">Frases R</td></tr>
<tr><td class="campo">Naturaleza de los riesgos espec�ficos atribuidos a las sustancias y preparados peligrosos</td></tr>
</table>
&nbsp;
<table class="tabla3" width="90%" align="center" height="100%" valign="middle">
<%=texto%>
</table>
<p align="center"><input type="button" class="boton" value="VER TODAS LAS FRASES R" onClick="location.href='busca_frases_r.asp?id=todas'"></p>
</body>
</html>