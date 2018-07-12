<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->
<!--#include file="dn_restringida.asp"-->

<%
	id = EliminaInyeccionSQL(request("id"))
	sql = "SELECT nombre FROM dn_risc_sustancias WHERE id="&id
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,objConnection,adOpenKeyset
	if not objRecordset.eof then nombresustancia=objRecordset("nombre")
	objRecordset.close
	set objRecordset=nothing
	cerrarconexion
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Enviar</title>
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
<link rel="stylesheet" type="text/css" href="estructura.css">
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<body>
<form action="dn_recomendar.asp?enviar=1&id=<%=id%>" method="post"><br />
<table class="tabla3" width="95%" align="center" height="100%" valign="middle" cellpadding="5">
<%
if request("enviar")<>1 then 'MOSTRAR FORMULARIO
%>
<tr>
  <td class=titulo3 colspan="2">Escriba su nombre y el email al que se enviar&aacute; el enlace a la ficha de la sustancia:</td>
  </tr>
 <tr>
  <td class=titulo3 align=right width="30%" >Nombre:</td>
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
    El email se utilizar&aacute; unicamente para enviar el enlace; los datos que nos facilite no ser&aacute;n almacenados. </td></tr>
	<tr>
  <td  align=center colspan="2" ><input type="submit" class="boton2" value="Enviar" /></td>
  </tr>
<%
else 'ENVIAR EMAIL
	response.write "<tr><td  align=center colspan='2' >"
	email=EliminaInyeccionSQL(request("email"))
	nombre=EliminaInyeccionSQL(request("nombre"))
	if email="" then
		errmes="Debe escribir un email.<br>"
	else
		if instr(email,".")=0 then  errmes="El email no es válido.<br> "
		if instr(email,"@")=0 then  errmes="El email no es válido.<br> "
	end if
	if nombre="" then errmes=errmes& "Debe escribir su nombre.<br> "
	if errmes<>"" then
		response.write errmes& "<br /><br /><input type='button' onclick='history.back()' value='volver' class='boton2' />"
	else

      enlace="http://217.13.81.22/istas/risctox/dn_risctox_ficha_sustancia.asp?id_sustancia="&id

      if (session("modo")="pruebas") then

        enlace=replace(enlace, "/web/", "/web/pruebas/web/")

      end if


			Set Mail = Server.CreateObject("Persits.MailSender")


			' Mail.Host = "smtp.istas.net"
			Mail.Host = "localhost"
			Mail.From = "inforisctox@istas.net"
			Mail.FromName = "ISTAS > risctox" ' Opcional
			' Mail.Username = "ncc550c"
			' Mail.Password = "***REMOVED**"

			Mail.AddAddress email
			Mail.Subject = Mail.EncodeHeader(nombre& " le recomienda que visite esta página")
			Mail.Body = nombre& " le recomienda que visite esta página para ver información toxicológica sobre la sustancia <strong>" &nombresustancia& "</strong>: <br><br><a href='" &enlace&"'>"&replace(enlace, "http://","")&"</a>"
			Mail.IsHTML = True
			On Error Resume Next
			Mail.Send      ' ó Mail.SendToQueue
			If Err <> 0 Then
	      			Response.Write "Ha ocurrido un error enviando el email. Por favor, vuelva a intentarlo en unos instantes."
			else
					response.write "El email se ha enviado correctamente. <br /><br /><input type='button' onclick='window.close()' value='cerrar ventana'  class='boton2' />"
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

