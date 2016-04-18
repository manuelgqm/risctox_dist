<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_conexion.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->
<!--#include file="../dn_restringida.asp"-->

<%
	id = EliminaInyeccionSQL(request("id"))
	sql = "SELECT nombre_ing FROM dn_risc_sustancias WHERE id="&id
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,objConnection,adOpenKeyset
	if not objRecordset.eof then nombresustancia=objRecordset("nombre_ing")
	objRecordset.close
	set objRecordset=nothing
	cerrarconexion
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>RISCTOX: Send</title>
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
<form action="dn_recomendar.asp?enviar=1&id=<%=id%>" method="post"><br />
<table class="tabla3" width="95%" align="center" height="100%" valign="middle" cellpadding="5">
<%
if request("enviar")<>1 then 'MOSTRAR FORMULARIO
%>
<tr>
  <td class=titulo3 colspan="2">Please enter your name and email in order to send a link to the substance page:</td>
  </tr>
 <tr>
  <td class=titulo3 align=right width="30%" >Name:</td>
  <td class=texto align=left  width="70%"><input type="text" name="nombre" /></td>
  </tr>
   <tr>
  <td class=titulo3 align=right >Email:<br /></td>
  <td class=texto align=left><input type="text" name="email" />
    </td>
  </tr>
<tr>
  <td class=titulo3 align=right ><br /></td>
  <td class=texto align=left>
    We use the email only to send you the link, we won't store it for any other purpose.</td></tr>
	<tr>
  <td  align=center colspan="2" ><input type="submit" class="boton2" value="Send" /></td>
  </tr>
<%
else 'ENVIAR EMAIL
	response.write "<tr><td  align=center colspan='2' >"
	email=EliminaInyeccionSQL(request("email"))
	nombre=EliminaInyeccionSQL(request("nombre"))
	if email="" then
		errmes="Debe escribir un email.<br>"
	else
		if instr(email,".")=0 then  errmes="The email is not valid.<br> "
		if instr(email,"@")=0 then  errmes="The email is not valid.<br> "
	end if
	if nombre="" then errmes=errmes& "Please enter your name.<br> "
	if errmes<>"" then
		response.write errmes& "<br /><br /><input type='button' onclick='history.back()' value='volver' class='boton2' />"
	else

      enlace="http://www.istas.net/risctox/en/dn_risctox_ficha_sustancia.asp?id_sustancia="&id

      if (session("modo")="pruebas") then

        enlace=replace(enlace, "/web/", "/web/pruebas/web/")

      end if


			Set Mail = Server.CreateObject("Persits.MailSender")


			' Mail.Host = "smtp.istas.net"
			Mail.Host = "localhost"
			Mail.From = "inforisctox@istas.net"
			Mail.FromName = "ISTAS > risctox" ' Opcional
			' Mail.Username = "ncc550c"
			' Mail.Password = "***REMOVED***"

			Mail.AddAddress email
			Mail.Subject = Mail.EncodeHeader(nombre& "  recommends that you visit this page")
			Mail.Body = nombre& "  recommends that you visit this page in order to view toxicologic information of substance <strong>" &nombresustancia& "</strong>: <br><br><a href='" &enlace&"'>"&replace(enlace, "http://","")&"</a>"
			Mail.IsHTML = True
			On Error Resume Next
			Mail.Send      '  Mail.SendToQueue
			If Err <> 0 Then
	      			Response.Write "An error has occurred. Please try again later."
			else
					response.write "Tha mail has been sent correctly. <br /><br /><input type='button' onclick='window.close()' value='Close window'  class='boton2' />"
			End If
	end if
	response.write "</td></tr>"
end if
%>
</table>
</form><br />
</body>
</html>
<%

%>

