<!--#include file="../../EliminaInyeccionSQL.asp"-->
 	<!--#include file="../dn_conexion.asp"-->
<%	


	tu_email = request("tu_email")
	tu_email = EliminaInyeccionSQL(tu_email)
	
	sql = "SELECT * FROM ECOINFORMAS_GENTE_NUEVO WHERE email='"&tu_email&"' and email<>''"
	
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	if not objRecordset.eof then
		
		clave = objRecordset("clave")
		contra = objRecordset("contra")
		email = objRecordset("email")

		   	   	        
		   	   	        
		   	   	        
	texto1 = "<p align=center><a href='http://www.istas.net/risctox/evalua/'>Evalua lo que usas</a></p>"
	texto2 = "Gracias por tu inter�s en la protecci�n del medio ambiente.<br>* Te recordamos que puedes darte de baja como usuario de ECOinformas en cualquier momento. Para m�s informaci�n: <a href='mailto:datospersonales@istas.net'>datos personales</a>."


		cuerpo = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd'>"
		cuerpo = cuerpo&"<HTML lang=es xmlns='http://www.w3.org/1999/xhtml'><HEAD><TITLE>ECOinformas:</TITLE>"&chr(13)
		cuerpo = cuerpo&"<LINK href='http://www.istas.net/ecoinformas/boletin/boletin.css' type=text/css rel=stylesheet></HEAD>"&chr(13)
		cuerpo = cuerpo&"<BODY bgColor=#ffffff>"&chr(13)
				'cuerpo = cuerpo&"<P>&nbsp;</P>"&chr(13)
		
		cuerpo = cuerpo & "<p class=texto align=left>Estimada se�ora, estimado se�or:</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left><b></b></p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Con la siguiente clave y contrase�a , tendr� acceso libre a toda la p�gina web y podr� aprovechar los servicios y productos gratuitos que le ofrecemos.</p>" & chr(13)
		cuerpo = cuerpo & "<p class='texto' align='center'>Clave: <b>"&clave&"</b><br>Contrase�a: <b>"&contra&"</b></p>" & chr(13)
		cuerpo = cuerpo & "<p class=texto align=left><b>�Sabe c�mo prevenir el riesgo qu�mico en su trabajo?</b></p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Entre en <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=575'><b>RISCTOX</b></a> con su clave y contrase�a e inf�rmese sobre las sustancias qu�micas que hay en su centro de trabajo y sobre las <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=576'><b>Alternativas</b></a> de sustiuci�n para proteger mejor su salud y el medio ambiente.</p>"&chr(13)
		'cuerpo = cuerpo & "<p class=texto align=left><a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=521'><b>�Conoce lo que usa?</b></a> Con esta gu�a onlina le ayudamos a conseguir informaci�n sobre las sustancias qu�micas presentes en su centro de trabajo.</p>" & chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>En la herramienta <a href='http://www.istas.net/risctox/evalua/'><b>Eval�a lo que usas</b></a> podr� conocer el riego que presentan los productos que utiliza en su empresa y compararlos con posibles alternativas.</p>"
		cuerpo = cuerpo & "<p class=texto align=left><b>�Quiere conocer la legislaci�n ambiental que afecta a su empresa?</b><br>"
		cuerpo = cuerpo & "<a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=710'><b>Legislaci�n on-line</b></a> le permitir� obtener informaci�n pr�ctica sobre la normativa ambiental aplicable a su empresa: las autorizaciones que debe obtener y la legislaci�n -estatal y auton�mica- que la afecta seg�n su actividad.</p>"
		cuerpo = cuerpo & "<p class=texto align=left>&nbsp;</p>"
		cuerpo = cuerpo & "<p class=texto align=left>* Le recordamos que puede darse de baja como usuario de ECOinformas en cualquier momento.</p>" & chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Los datos que nos facilita ser�n incorporados a un fichero titularidad de ISTAS. La finalidad del tratamiento de sus datos la constituye la posibilidad de difusi�n por correo electr�nico y ordinario de informaci�n y materiales de ECOinformas; la promoci�n de la salud laboral y la protecci�n del medio ambiente a trav�s de la remisi�n de informaci�n sobre nuestros productos editoriales y actividades; auditor�a por parte de la Fundaci�n Biodiversidad que se compromete a su vez a cumplir la Ley Org�nica de Protecci�n de Datos de car�cter Personal (LOPD).</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Para darse de baja de nuestra base de datos puede enviar una notificaci�n a la direcci�n de correo electr�nico <a href=mailto:datospersonales@istas.net?subject=Baja>datospersonales@istas.net</a> con la refencia 'Baja Evalua lo que usas' o por correo ordinario a la direcci�n de ISTAS: Calle General Cabrera, 21. 28020 Madrid.</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Finalmente puede ejercer sus derechos de acceso, rectificaci�n o cancelaci�n en cumplimiento de lo establecido en la LOPD, enviando una solicitud por escrito, acompa�ada de una copia de su D.N.I. indicando como referencia 'Protecci�n de datos' dirigida a ISTAS con domicilio sito en la Calle General Cabrera, 21, 28020 Madrid.</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Para m�s informaci�n: <a href=mailto:datospersonales@istas.net>datospersonales@istas.net</a></p>"&chr(13)
		
				
		cuerpo = cuerpo&"<P>&nbsp;</P>"&chr(13)

		
		cuerpo = cuerpo & "</BODY></HTML>"
		
		

		if email<>"" then
			
			Set Mail = Server.CreateObject("Persits.MailSender")
			Mail.Host = "localhost"
			'Mail.Port = 587
			Mail.From = "xip@istas.net"
			Mail.FromName = "ISTAS: Riesgo qu�mico: Evalua lo que usas" ' Opcional 
			Mail.AddAddress email
			Mail.Subject = Mail.EncodeHeader("Recordatorio de acceso a Evalua lo que usas")
			Mail.Body = cuerpo
			Mail.IsHTML = True
			On Error Resume Next
			Mail.Send      ' � Mail.SendToQueue
			If Err <> 0 Then
	      			Response.Write "Error en la cuenta " & email & ": " & Err.Description & "<br>" 
			End If 
		end if

	else
		acceso = "no"
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

'Sergio->redirecciono a acceso, no siguo para abajo

response.redirect "acceso.asp?recordar=1"
response.End()

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: Base de datos de sustancias t�xicas y peligrosas RISCTOX</title>
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

</head>
<body>
<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
			<div id="encabezado_nuevo1">
			<table width="100%" cellpadding=0 border=0>
			<tr><td width="215" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="142" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="166" height="78" onclick="location.href='index.asp?idpagina=549'" style="cursor:hand">&nbsp;</td>
			    <td width="160" height="78" onclick="location.href='index.asp?idpagina=550'" style="cursor:hand">&nbsp;</td>
			    <td width="25"  height="78" align="center">
			    	<a href="mailto:ppedroso@istas.ccoo.es?subject=Contacto ECOinformas"><img src="imagenes/ico_contacto.gif" border="0" alt="Contacto"></a><br>
			    	<a href="busqueda.asp"><img src="imagenes/ico_busqueda.gif" border="0" alt="busqueda"></a><br>
			    	<a href="index.asp?idpagina=560"><img src="imagenes/ico_ayuda.gif" border="0" alt="ayuda"></a>
			    </td>
			</tr>
			</table>
			</div>
			<div id="menusup1">
			<table border="0" cellspacing="5" cellpadding="0">
            			<tr class="textmenusup"><td class=textmenusup>P&aacute;gina de identificaci�n</td>
          		</table>
			</div>
			
			<div class="textsubmenu" id="submenusup">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
            			<tr>
              				<td width="100%" valign="top">Est&aacute;s en: identificaci&oacute;n para acceso a zonas restringidas</td>
            			</tr>
          		</table>
			</div>
			
			<div id="cuerpo">
					<p>&nbsp;</p>
					<center>
		    			<% if acceso="si" then %>
					<div class="tabla" width="90%">
					<p class="texto">Acabamos de enviarte un correo electr�nico a tu direcci�n con la clave y contrase�a para que puedas acceder libremente a todo el espacio web de ECOinformas.</p>
					</div>
					
					<p>&nbsp;</p>

					<div id="identifica">
					<form name="form1" id="form1" method="post" action="identifica.asp">
					<input type="hidden" name="idpagina" value="<%=request("idpagina")%>">
					<table width="100%" border="0" cellspacing="2" cellpadding="2" align="center" bgcolor="#00AC5A">
				  		<tr bgcolor="#006600">
                					<td colspan="2">Identificaci&oacute;n</td>
                				</tr>
              					<tr>
                					<td>Clave:</td>
                					<td><input name="clave" type="text" class="campoform" id="clave" size="8" maxlenght="20" /></td>
              					</tr>
              					<tr>
                					<td>Contrase&ntilde;a:</td>
                					<td><input name="contra" type="password" class="campoform" id="contra" size="8" maxlenght="20" /></td>
              					</tr>
              					<tr>
                					<td>&nbsp;</td>
                					<td><input class="boton" type="submit" name="Submit" value="Enviar" /></td>
              					</tr>
              				</table>
				  	</form>
            </div>
					</center>
					<br>&nbsp;
				<% else %>
					
					<form name="form_recordar" action="recordar.asp" method="POST">
					<p class="textoc">Lo sentimos. Tu e-mail no coincide con ninguno de los que est�n registrados.
						 Si quieres volver a escribir tu email para comprobarlo de nuevo puedes hacerlo.</p>
					<p class="textoc">Tu e-mail:&nbsp;<input type="text" class="campo" size="50" maxlenght="200" name="email">&nbsp;<input class="boton" type="submit" name="Submit" value="Enviar" /></p>
				  </form>

		    			<p>&nbsp;</p>
							
					<table width="80%" border="0" cellspacing="2" cellpadding="2" align="center" class="tabla">
					<tr><td class="textoc">Te aconsejamos que te registres de nuevo para recibir inmediatamente una nueva clave y contrase�a.</td></tr>
					<tr><td class="textoc"><input type="button" class="boton" value="solicitar acceso libre" onclick="location.href='formulario2.asp'"/></td></tr></table>
				<% end if %>
		    			<p>&nbsp;</p>
		    			<p>&nbsp;</p>

     			<map name="Map1" id="Map1">
            		<area shape="rect" coords="300,18,392,80" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
            		<area shape="rect" coords="393,18,539,80" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,18,694,80" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<img src="imagenes/pie1.jpg" width="708" border="0" usemap="#Map1">
      			
    			
    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>
</body>
</html>
