<!--#include file="../../EliminaInyeccionSQL.asp"-->
<!--#include file="../dn_conexion.asp"-->
<%

email = eliminainyeccionSQL(unquote(request("email")))

Sub StrRandomize(strSeed)
	Dim i, nSeed
	
	nSeed = CLng(0)
	For i = 1 To Len(strSeed)
	nSeed = nSeed Xor ((256 * ((i - 1) Mod 4) * AscB(Mid(strSeed, i, 1))))
	Next
	
	'Randomizar
	Randomize nSeed
End Sub

Function GeneratePassword(nLength)
	Dim i, bMadeConsonant, c, nRnd
	
	Const strDoubleConsonants = "bdfglmnpst"
	Const strConsonants = "bcdfghklmnpqrstv"
	Const strVocal = "1234567890"
	
	GeneratePassword = ""
	bMadeConsonant = False
	
	For i = 0 To nLength
	'Obtener un número aleatorio ente 0 y 1
	nRnd = Rnd
	
	If GeneratePassword <> "" AND (bMadeConsonant <> True) AND (nRnd < 0.15) Then
	'double consonante
	c = Mid(strDoubleConsonants, Int(Len(strDoubleConsonants) * Rnd + 1), 1)
	'response.write int(Len(strDoubleConsonants) * Rnd + 1)
	'response.write "<br>"
	c = c & c
	i = i + 1
	bMadeConsonant = True
	Else
	
	
	If (bMadeConsonant <> True) And (nRnd < 0.95) Then
	'Simple consonant
	c = Mid(strConsonants, Int(Len(strConsonants) * Rnd + 1), 1)
	bMadeConsonant = True
	
	Else
	'Si la útima letra fué una consonate, crear una vocal
	c = Mid(strVocal,Int(Len(strVocal) * Rnd + 1), 1)
	bMadeConsonant = False
	End If
	End If
	'Sumar Letra
	GeneratePassword = GeneratePassword & c
	Next
	
	'Es el password demasiado corto o demasiado largo?
	If Len(GeneratePassword) > nLength Then
	GeneratePassword = Left(GeneratePassword, nLength)
	End If
End Function



StrRandomize CStr(Now) & CStr(Rnd)
contra=GeneratePassword(5)

'Primero compruebo si el correo está en la BBDD
'si está le envío la contraseña al correo
sql = "select * from ECOINFORMAS_GENTE_NUEVO where email='"&email&"'"
Set objR = OBJConnection.Execute(sql)
if objr.eof then'Si no existe el correo, creamos el usuario
	orden = "Insert into ECOINFORMAS_GENTE_NUEVO(email, clave, contra) values('"&email&"','"&email&"','"&contra&"')"
	Set objRecordset = OBJConnection.Execute(orden)
	orden = "SELECT max(idgente) as ultimo FROM ECOINFORMAS_GENTE_NUEVO"
	Set objRecordset = OBJConnection.Execute(orden)
	ultimo = objRecordset("ultimo")
	session("id_ecogente2") = cstr(ultimo)
else 'si ya existe le enviamos el correo
	ya_existe=1
	clave = objr("clave")
	contra = objr("contra")

end if



		cuerpo = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd'>"
		cuerpo = cuerpo&"<HTML lang=es xmlns='http://www.w3.org/1999/xhtml'><HEAD><TITLE>ECOinformas:</TITLE>"&chr(13)
		cuerpo = cuerpo&"<LINK href='http://www.istas.net/risctox/evalua/boletin.css' type=text/css rel=stylesheet></HEAD>"&chr(13)
		cuerpo = cuerpo&"<BODY bgColor=#ffffff>"&chr(13)

		cuerpo = cuerpo & "<p class=texto align=left>Estimada señora, estimado señor:</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left><b>Gracias por registrarse como usuario de Evalua lo que usas.</b></p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Con la clave y contraseña que le asignamos a continuación, tendrá acceso libre a toda la página web y podrá aprovechar los servicios y productos gratuitos que le ofrecemos.</p>" & chr(13)
		cuerpo = cuerpo & "<p class='texto' align='center'>Clave: <b>"&email&"</b><br>Contraseña: <b>"&contra&"</b></p>" & chr(13)
		cuerpo = cuerpo & "<p class=texto align=left><b>¿Sabe cómo prevenir el riesgo químico en su trabajo?</b></p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Entre en <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=575'><b>RISCTOX</b></a> con su clave y contraseña e infórmese sobre las sustancias químicas que hay en su centro de trabajo y sobre las <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=576'><b>Alternativas</b></a> de sustiución para proteger mejor su salud y el medio ambiente.</p>"&chr(13)
		'cuerpo = cuerpo & "<p class=texto align=left><a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=521'><b>¿Conoce lo que usa?</b></a> Con esta guía onlina le ayudamos a conseguir información sobre las sustancias químicas presentes en su centro de trabajo.</p>" & chr(10)& chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>En la herramienta <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=961'><b>Evalúa lo que usas</b></a> podrá conocer el riego que presentan los productos que utiliza en su empresa y compararlos con posibles alternativas.</p>"
		
		cuerpo = cuerpo & "<p class=texto align=left>Para darse de baja de nuestra base de datos puede enviar una notificación a la dirección de correo electrónico <a href=mailto:datospersonales@istas.net?subject=Baja>datospersonales@istas.net</a> con la refencia 'Baja Evalua lo que usas' o por correo ordinario a la dirección de ISTAS: Calle General Cabrera, 21. 28020 Madrid.</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Finalmente puede ejercer sus derechos de acceso, rectificación o cancelación en cumplimiento de lo establecido en la LOPD, enviando una solicitud por escrito, acompañada de una copia de su D.N.I. indicando como referencia 'Protección de datos' dirigida a ISTAS con domicilio sito en la Calle General Cabrera, 21, 28020 Madrid.</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Para más información: <a href=mailto:datospersonales@istas.net>datospersonales@istas.net</a></p>"&chr(13)
		
		'cuerpo = cuerpo & "<p class=texto align=left>" & texto1 & "</p>" & chr(13)
		'cuerpo = cuerpo & "<p class=texto align=left>" & texto2 & "</p>" & chr(10)& chr(13)
		
		
		cuerpo = cuerpo & "</BODY></HTML>"

		if email<>"" then
			Set Mail = Server.CreateObject("Persits.MailSender")
			Mail.Host = "localhost"
			'Mail.Port = 587
			Mail.From = "xip@istas.net"
			Mail.FromName = "Evalua lo que usas" ' Opcional 
			'Mail.Username = "say5151"
			'Mail.Password = "***REMOVED***"
			Mail.AddAddress email
			Mail.Subject = Mail.EncodeHeader("Recordatorio de acceso a Evalua lo que usas")
			Mail.Body = cuerpo
			Mail.IsHTML = True
			On Error Resume Next
			Mail.Send      ' ó Mail.SendToQueue
			If Err <> 0 Then
	      			Response.Write "Error en la cuenta " & email & ": " & Err.Description & "<br>" 
			End If 
		end if


if ya_existe=1 then 'le redirecciono y que acceda, no le muestro la clave en la 
		response.redirect("./acceso.asp")
		response.end()
end if





FUNCTION unQuote(s)
  pos = Instr(s, "'")
  While pos > 0 
    s = Mid(s,1,pos) & "'" & Mid(s,pos+1)
    pos = InStr(pos+2, s, "'")
  Wend
  pos = Instr(s, """") 
  While pos > 0 
    s = Mid(s,1,pos-1) & "''" & Mid(s,pos+1)
    pos = InStr(pos+2, s, """")
  Wend
  unQuote = Trim(s)
END FUNCTION




%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: BBDD Evalúa lo que usas</title>
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
<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
			
			<!--#include file="../dn_cabecera.asp"-->
			<div id="texto">
				<div class="texto">
				
				<br>&nbsp;
				<table width="90%" align="center" clas=tabla>
				<tr><td class="texto" align="center">Gracias por remitir tus datos. Tus claves de acceso son:
				<p class="texto" width="50%">CLAVE:&nbsp;<b><%=email%></b></p>
				<p class="texto" width="50%">CONTRASEÑA:&nbsp;<b><%=contra%></b></p>
				Acabamos de enviarte un correo a la dirección <%=email%> con tus claves de acceso para que puedas volver a entrar en todas las herramientas de Evalua lo que usas.<br>
				</td></tr></table>
				
				
				<br>&nbsp;
				<p class="texto" width="100%" align="center">			
				<input type=button class=boton value="Ir a Herramienta de Autoevaluación" onclick=location.href="dn_auto_herramienta.asp">
                </p>
				
				<br>&nbsp;
				
				<table width="90%" align="center">
				<tr><td class=texto><%=texto2%></td></tr>
				<tr><td class="texto">&nbsp;</td></tr>
				</table>
				
				</div>
				
				<p><br /><br />&nbsp;</p>
			</div>

			<img src="../imagenes/pie_risctox.gif" width="708" border="0">

    			</div>
    		</div>
		<div id="sombra_abajo"></div>
	</div>
</body>
</html>