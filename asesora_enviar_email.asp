<!--#include file="EliminaInyeccionSQL.asp"-->
<%

function enviar_mail()


        if cstr(modelo)="0" or cstr(modelo)="3" then
	 orden = "SELECT nombre,apellidos,sexo,email,clave,contra FROM ECOINFORMAS_GENTE WHERE idgente="&idgente
	 Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	 Set objRecordset = OBJConnection.Execute(orden)
	 nombre = objRecordset("nombre")&" "&objRecordset("apellidos")
	 sexo = "o"
	 if objRecordset("sexo")=75 then sexo="a"
	 destinatarios = trim(objRecordset("email"))
	end if 
	
	cuerpo = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd'>"
	cuerpo = cuerpo&"<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd'>"
	cuerpo = cuerpo&"<HTML lang=es xmlns='http://www.w3.org/1999/xhtml'><HEAD><TITLE>ECOinformas:</TITLE>"
	cuerpo = cuerpo&"<META http-equiv=Content-Type content='text/html; charset=iso-8859-1'>"
	cuerpo = cuerpo&"<META content=ECOinformas name=Title>"
	cuerpo = cuerpo&"<META content='XiP multimèdia' name=Author>"
	cuerpo = cuerpo&"<META content='Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME' name=description>"
	cuerpo = cuerpo&"<META content='Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME' name=Subject>"
	cuerpo = cuerpo&"<META content='Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME' name=Keywords>"
	cuerpo = cuerpo&"<META content=Spanish name=Language>"
	cuerpo = cuerpo&"<META content='15 days' name=Revisit>"
	cuerpo = cuerpo&"<META content=Global name=Distribution>"
	cuerpo = cuerpo&"<META content=All name=Robots><LINK media=screen href='http://www.istas.net/ecoinformas/boletin/boletin.css' type=text/css rel=stylesheet>"
	cuerpo = cuerpo&"<META content='MSHTML 6.00.2800.1505' name=GENERATOR></HEAD>"
	cuerpo = cuerpo&"<BODY bgColor=#ffffff>"
	cuerpo = cuerpo&"<DIV>&nbsp;</DIV>"
	cuerpo = cuerpo&"<DIV id=contenedor>"
	cuerpo = cuerpo&"<DIV id=sombra_arriba></DIV>"
	cuerpo = cuerpo&"<DIV id=sombra_lateral>"
	cuerpo = cuerpo&"<DIV id=caja>"
	cuerpo = cuerpo&"<DIV id=encabezado></DIV>"
	cuerpo = cuerpo&"<DIV id=menusup><SPAN class=texto align=right>Madrid, "&formatdatetime(now(),1)&"</SPAN></DIV>"
	cuerpo = cuerpo&"<DIV class=textsubmenu id=submenusup></DIV></DIV>"
	cuerpo = cuerpo&"<DIV id=cuerpo>"
	cuerpo = cuerpo&"<P>&nbsp;</P>"
	
	' Modelo que reciba el usuario final al hacer su pregunta de asesoramiento
	if cstr(modelo)="0" then
	 cuerpo = cuerpo&"<P class=texto align=left><b>Estimad"&sexo&"&nbsp;"&nombre&":</b></P>"
	 cuerpo = cuerpo&"<P class=texto align=left>Este es un mensaje automático.<br>"
	 cuerpo = cuerpo&"<br><br>"
	 cuerpo = cuerpo&"<P class=texto align=left>El equipo de ECOinformas tratará tu asunto y será atendido a la mayor brevedad."
	 cuerpo = cuerpo&"<br><br>"
	 cuerpo = cuerpo&"<P class=texto align=left>Gracias por tu interés en la promoción de salud laboral y protección del medio ambiente.</P>"
	 cuerpo = cuerpo&"<P class=texto align=left><b>Equipo ECOinformas</b></P>"
	end if

        ' Modelo que reciba el coordinador cuando se hace un asesoramiento
	if cstr(modelo)="1" then
	 cuerpo = cuerpo&"<P class=texto align=left>Recibida solicitud de asesoramiento <a href='http://www.ecoinformas.com'>http://www.ecoinformas.com</a>.<br>"
	 cuerpo = cuerpo&"<br><br>"
	 
	end if
	
	' Modelo que recibe el asesor cuando se le asigna un asesoramiento
	if cstr(modelo)="2" then
	 cuerpo = cuerpo&"<P class=texto align=left>Tienes asignada una solicitud de asesoramiento <a href='http://www.ecoinformas.com'>http://www.ecoinformas.com</a>."
	 cuerpo = cuerpo&"<br><br>"
	end if

	' Modelo que reciba el usuario final cuando el asesor le responde a su pregunta
	if cstr(modelo)="3" then
	 cuerpo = cuerpo&"<P class=texto align=left><b>Estimad"&sexo&"&nbsp;"&nombre&":</b></P>"
	 cuerpo = cuerpo&"<P class=texto align=left>Uno(a) de nuestros(as) técnicos(as) tiene respuesta a tu solicitud de asesoramiento.<br>"
	 cuerpo = cuerpo&"<br><br>"
	 cuerpo = cuerpo&"<P class=texto align=left>Puedes consultarla en la página web <a href='http://www.ecoinformas.com'>http://www.ecoinformas.com</a>."
	 cuerpo = cuerpo&"<br><br>"
	 cuerpo = cuerpo&"<P class=texto align=left>Gracias por tu interés en la promoción salud laboral y protección del medio ambiente.</P>"
	 cuerpo = cuerpo&"<P class=texto align=left><b>Equipo ECOinformas</b></P>"
	end if

	cuerpo = cuerpo&"<MAP id=Map name=Map>"
	cuerpo = cuerpo&"<AREA shape=RECT target=_blank alt='Fundación Biodiversidad' coords=267,35,339,84 href='http://www.fundacion-biodiversidad.es'>"
	cuerpo = cuerpo&"<AREA shape=RECT target=_blank alt='Instituto Sindical de Trabajo, Ambiente y Salud' coords=342,34,471,84 href='http://www.istas.ccoo.es'>"
	cuerpo = cuerpo&"<AREA shape=RECT target=_blank alt='Fondo Social Europeo' coords=472,35,599,82 href='http://www.mtas.es/UAFSE/default.htm'>"
	cuerpo = cuerpo&"</MAP>"
	cuerpo = cuerpo&"<DIV><IMG src='http://www.istas.net/ecoinformas/boletin/pie.jpg' useMap=#Map border=0></DIV></DIV>"
	cuerpo = cuerpo&"<DIV>"
	cuerpo = cuerpo&"<DIV id=sombra_abajo></DIV>"
	cuerpo = cuerpo&"</DIV></DIV></DIV>"
	cuerpo = cuerpo&"</BODY></HTML>"


asunto = "Asesoramiento ECOinformas"

enviar_mail = false
if destinatarios<>"" then
 correos = Split(destinatarios,",") 
 direcciones = ""
 total_destinatarios = UBound(correos)
 for i=0 to total_destinatarios
	if correos(i)<>"" then
	Set Mail = Server.CreateObject("Persits.MailSender")
	Mail.Host = "smtp.ecoinformas.com"
	Mail.From = "jdejong@istas.net"
	Mail.FromName = "ECOinformas" ' Opcional 
	Mail.Username = "say5151"
	Mail.Password = "***REMOVED**"
	Mail.AddAddress correos(i)
	Mail.Subject = Mail.EncodeHeader(asunto)
	Mail.Body = cuerpo
	Mail.IsHTML = True
	On Error Resume Next
	Mail.Send      ' ó Mail.SendToQueue
	If Err <> 0 Then
		Response.Write "Error en la cuenta " & email_dest & ": " & Err.Description & "<br>" 
	else
	 enviar_mail = true	
	End If 
	end if
 next	
end if

end function

%>