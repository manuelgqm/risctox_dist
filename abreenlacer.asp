<%
	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	
 	if session("id_ecogente")="" then response.redirect "acceso.asp?idenlace="&request("idenlace")
			

idenlace = request("idenlace")

sql = "SELECT visitas,url,afiliacion FROM ENL_ENLACES WHERE id="&idenlace
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
set objRecordset = OBJConnection.Execute(sql)

visitas = objRecordset("visitas")
url = objRecordset("url")
afiliacion = objRecordset("afiliacion")

sql = "UPDATE ENL_ENLACES SET visitas="&visitas+1&" WHERE id="&idenlace
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
set objRecordset = OBJConnection.Execute(sql)

IP = Request.ServerVariables("REMOTE_ADDR")
Set MiBrowser = Server.CreateObject("MSWC.BrowserType")
navegador = MiBrowser.Browser
orden = "INSERT INTO ENL_VISITAS (fecha,hora,IP,navegador,idenlace,idgente) VALUES ('"&date()&"','"&time()&"','"&IP&"','"&navegador&"','"&idenlace&"','"&session("id_ecogente")&"')"
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
Set objRecordset = OBJConnection.Execute(orden)

%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: abre enlace</title>
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

<frameset framespacing="0" border="0" rows="40,*" frameborder="0">
   <frame name="enlace_superior" src="cabecera_enlace.asp?url=Enlace restringido para usuarios registrados"  scrolling="no">
   <frame name="enlace_inferior" src="<%=url%>" class="tabla">
</frameset>
</html>
