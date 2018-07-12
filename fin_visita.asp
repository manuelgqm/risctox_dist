<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

	if cstr(session("id_ecogente"))="4" then 
		session("id_ecogente")=""
		response.redirect "http://www.ecoinformas.com/"
	end if

	sql = "SELECT * FROM ECOINFORMAS_GENTE WHERE idgente="&session("id_ecogente")
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	if not objRecordset.eof then
		usuario = objRecordset("nombre")&" "&objRecordset("apellidos")
		usuario_sexo = "o"
		if objRecordset("sexo")=75 then usuario_sexo = "a"
		usuario_texto = "Usuari" & usuario_sexo & " identificad" & usuario_sexo & ":&nbsp;" & usuario & "&nbsp;"
		email = cstr(objRecordset("email"))
		clave = cstr(objRecordset("clave"))
		contra = cstr(objRecordset("contra"))
		
		SET04 = (objRecordset("SET04"))
		EGP01 = (objRecordset("EGP01"))
		EGP02 = (objRecordset("EGP02"))
		EGP03 = (objRecordset("EGP03"))
		EGP04 = (objRecordset("EGP04"))
		EGP05 = (objRecordset("EGP05"))
		EGP06 = (objRecordset("EGP06"))
		EGP07 = (objRecordset("EGP07"))
		AE01 = (objRecordset("AE01"))
		AE02 = (objRecordset("AE02"))
		AE03 = (objRecordset("AE03"))
		AE04 = (objRecordset("AE04"))
		AE05 = (objRecordset("AE05"))
		AE06 = (objRecordset("AE06"))
		SEP01 = (objRecordset("SEP01"))
		SEP02 = (objRecordset("SEP02"))
		SEP03 = (objRecordset("SEP03"))
	end if
		   	   	        
		   	   	        
		   	   	        
if (SET04="1") or (EGP01="1") or (EGP02="1") or (EGP03="1") or (EGP04="1") or (EGP05="1") or (EGP06="1") or (EGP07="1") or (AE01="1") or (AE02="1") or (AE03="1") or (AE04="1") or (AE05="1") or (AE06="1") or (SEP01="1") or (SEP02="1") or (SEP03="1")  then
	texto1 = "Si has solicitado el envío de materiales impresos por correo postal, te avisamos que antes de poder enviartelos tenemos que valorar si perteneces a los colectivos a los que va dirigido ECOinformas. Si es el caso, te avisaremos sobre el envío. Si no pertenecieras a los colectivos elegibles, no recibirás las publicaciones impresas, pero sí podrás descargar los archivos en PDF desde <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=548'>www.ecoinformas.com</a>"
	texto2 = "Gracias por tu interés en la protección del medio ambiente.<br>* Te recordamos que puedes darte de baja como usuario de ECOinformas en cualquier momento. Para más información: <a href='mailto:datospersonales@istas.net'>datos personales</a>."
else
	texto1 = "<p align=center><a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=548'>www.ecoinformas.com</a></p>"
	texto2 = "Gracias por tu interés en la protección del medio ambiente.<br>* Te recordamos que puedes darte de baja como usuario de ECOinformas en cualquier momento. Para más información: <a href='mailto:datospersonales@istas.net'>datos personales</a>."
end if

if 1=0 then
		cuerpo = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd'>"
		'cuerpo = cuerpo&"<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd'>"
		cuerpo = cuerpo&"<HTML lang=es xmlns='http://www.w3.org/1999/xhtml'><HEAD><TITLE>ECOinformas:</TITLE>"&chr(13)
		'cuerpo = cuerpo&"<META http-equiv=Content-Type content='text/html; charset=iso-8859-1'>"&chr(13)
		'cuerpo = cuerpo&"<META content=ECOinformas name=Title>"&chr(13)
		'cuerpo = cuerpo&"<META content='XiP multimèdia' name=Author>"&chr(13)
		'cuerpo = cuerpo&"<META content='Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME' name=description>"&chr(13)
		'cuerpo = cuerpo&"<META content='Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME' name=Subject>"&chr(13)
		'cuerpo = cuerpo&"<META content='Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME' name=Keywords>"&chr(13)
		'cuerpo = cuerpo&"<META content=Spanish name=Language>"&chr(13)
		'cuerpo = cuerpo&"<META content='15 days' name=Revisit>"&chr(13)
		'cuerpo = cuerpo&"<META content=Global name=Distribution>"&chr(13)
		'cuerpo = cuerpo&"<META content=All name=Robots>
		'cuerpo = cuerpo&"<META content='MSHTML 6.00.2800.1505' name=GENERATOR>"&chr(13)
		cuerpo = cuerpo&"<LINK href='http://www.istas.net/ecoinformas/boletin/boletin.css' type=text/css rel=stylesheet></HEAD>"&chr(13)
		cuerpo = cuerpo&"<BODY bgColor=#ffffff>"&chr(13)
		'cuerpo = cuerpo&"<DIV>&nbsp;</DIV>"&chr(13)
		cuerpo = cuerpo&"<DIV id=contenedor>"&chr(13)
		cuerpo = cuerpo&"<DIV id=sombra_arriba></DIV>"&chr(13)
		cuerpo = cuerpo&"<DIV id=sombra_lateral>"&chr(13)
		cuerpo = cuerpo&"<DIV id=caja>"&chr(13)
		cuerpo = cuerpo&"<DIV id=encabezado></DIV>"&chr(13)
		cuerpo = cuerpo&"<DIV id=menusup><SPAN class=texto align=right>Madrid, "&formatdatetime(now(),1)&"</SPAN></DIV>"&chr(13)
		cuerpo = cuerpo&"<DIV class=textsubmenu id=submenusup></DIV></DIV>"&chr(13)
		cuerpo = cuerpo&"<DIV id=cuerpo>"&chr(13)
		'cuerpo = cuerpo&"<P>&nbsp;</P>"&chr(13)
		
		cuerpo = cuerpo & "<p class=texto align=left>Estimada señora, estimado señor:</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left><b>Gracias por registrarse como usuario de ECOInformas*.</b></p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Con la clave y contraseña que le asignamos a continuación, tendrá acceso libre a toda la página web y podrá aprovechar los servicios y productos gratuitos que le ofrecemos.</p>" & chr(13)
		cuerpo = cuerpo & "<p class='texto' align='center'>Clave: <b>"&clave&"</b><br>Contraseña: <b>"&contra&"</b></p>" & chr(13)
		cuerpo = cuerpo & "<p class=texto align=left><b>¿Sabe cómo prevenir el riesgo químico en su trabajo?</b></p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Entre en <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=575'><b>RISCTOX</b></a> con su clave y contraseña e infórmese sobre las sustancias químicas que hay en su centro de trabajo y sobre las <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=576'><b>Alternativas</b></a> de sustiución para proteger mejor su salud y el medio ambiente.</p>"&chr(13)
		'cuerpo = cuerpo & "<p class=texto align=left><a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=521'><b>¿Conoce lo que usa?</b></a> Con esta guía onlina le ayudamos a conseguir información sobre las sustancias químicas presentes en su centro de trabajo.</p>" & chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>En la herramienta <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=961'><b>Evalúa lo que usas</b></a> podrá conocer el riego que presentan los productos que utiliza en su empresa y compararlos con posibles alternativas.</p>"
		cuerpo = cuerpo & "<p class=texto align=left><b>¿Quiere conocer la legislación ambiental que afecta a su empresa?</b><br>"
		cuerpo = cuerpo & "<a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=710'><b>Legislación on-line</b></a> le permitirá obtener información práctica sobre la normativa ambiental aplicable a su empresa: las autorizaciones que debe obtener y la legislación -estatal y autonómica- que la afecta según su actividad.</p>"
		'cuerpo = cuerpo & "<p class=texto align=left><b>Las personas que con su clave y contraseña entren en las páginas ¿Conoces lo que usas?, RISCTOX y Alternativas, recibirán por correo postal un vídeo sobre la prevención de riesgo químico gratuito.</b></p>"&chr(13)
		'cuerpo = cuerpo & "<center><object classid='clsid:d27cdb6e-ae6d-11cf-96b8-444553540000' codebase='http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0' width=530 height=100 id='anima0' align='middle'><param name='allowScriptAccess' value='sameDomain' /><param name='movie' value='http://www.istas.net/recursos/ANI/ISTAS_01076.swf' /><param name='quality' value='high' /><param name='wmode' value='transparent' /><param name='bgcolor' value='#ffffff' /><embed src='http://www.istas.net/recursos/ANI/ISTAS_01076.swf' quality='high' wmode='transparent' bgcolor='#ffffff' width=530 height=100 name='' align='middle' allowScriptAccess='sameDomain' type='application/x-shockwave-flash' pluginspage='http://www.macromedia.com/go/getflashplayer' /></object></center>"&chr(13)
		'cuerpo = cuerpo & "<p class=texto align=center><a href='http://www.ecoinformas.com/inicio.asp'><img src='http://www.istas.net/ecoinformas/web/imagenes/ISTAS_01076.gif' border=0></a></p>"
		'cuerpo = cuerpo & "<p class=texto align=left><b>Gracias por su interés en la protección del medio ambiente.</b></p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>&nbsp;</p>"
		cuerpo = cuerpo & "<p class=texto align=left>* Le recordamos que puede darse de baja como usuario de ECOinformas en cualquier momento.</p>" & chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Los datos que nos facilita serán incorporados a un fichero titularidad de ISTAS. La finalidad del tratamiento de sus datos la constituye la posibilidad de difusión por correo electrónico y ordinario de información y materiales de ECOinformas; la promoción de la salud laboral y la protección del medio ambiente a través de la remisión de información sobre nuestros productos editoriales y actividades; auditoría por parte de la Fundación Biodiversidad que se compromete a su vez a cumplir la Ley Orgánica de Protección de Datos de carácter Personal (LOPD).</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Para darse de baja de nuestra base de datos puede enviar una notificación a la dirección de correo electrónico <a href=mailto:datospersonales@istas.net?subject=Baja>datospersonales@istas.net</a> con la refencia 'Baja ECOinformas' o por correo ordinario a la dirección de ISTAS: Calle General Cabrera, 21. 28020 Madrid.</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Finalmente puede ejercer sus derechos de acceso, rectificación o cancelación en cumplimiento de lo establecido en la LOPD, enviando una solicitud por escrito, acompañada de una copia de su D.N.I. indicando como referencia 'Protección de datos' dirigida a ISTAS con domicilio sito en la Calle General Cabrera, 21, 28020 Madrid.</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Para más información: <a href=mailto:datospersonales@istas.net>datospersonales@istas.net</a></p>"&chr(13)
		
		'cuerpo = cuerpo & "<p class=texto align=left>" & texto1 & "</p>" & chr(13)
		'cuerpo = cuerpo & "<p class=texto align=left>" & texto2 & "</p>" & chr(13)
		
		cuerpo = cuerpo & "<P>&nbsp;</P>"&chr(13)

		cuerpo = cuerpo & "<p align=center><map name='mapa_poste'><area href='http://www.ecoinformas.com' shape='rect' coords='41, 21, 195, 66'><area href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=575' shape='rect' coords='87, 79, 233, 119'><area href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=576' shape='rect' coords='19, 130, 142, 168'><area href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=961' shape='rect' coords='88, 174, 233, 216'><area href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=710' shape='rect' coords='16, 227, 149, 263'><area href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=566' shape='rect' coords='99, 268, 229, 311'></map>"&chr(13)
		cuerpo = cuerpo & "<img border='0' src='http://www.istas.net/ecoinformas/web/imagenes/eco_poste.gif' usemap='#mapa_poste' width='242' height='362'></p>"
		cuerpo = cuerpo & "<P>&nbsp;</P>"&chr(13)
		
		cuerpo = cuerpo & "<MAP id=Map name=Map>"&chr(13)
		cuerpo = cuerpo & "<AREA shape=RECT target=_blank alt='Fundación Biodiversidad' coords=267,35,339,84 href='http://www.fundacion-biodiversidad.es'>"&chr(13)
		cuerpo = cuerpo & "<AREA shape=RECT target=_blank alt='Instituto Sindical de Trabajo, Ambiente y Salud' coords=342,34,471,84 href='http://www.istas.ccoo.es'>"&chr(13)
		cuerpo = cuerpo & "<AREA shape=RECT target=_blank alt='Fondo Social Europeo' coords=472,35,599,82 href='http://www.mtas.es/UAFSE/default.htm'>"&chr(13)
		cuerpo = cuerpo & "</MAP>"&chr(13)
		cuerpo = cuerpo & "<DIV><IMG src='http://www.istas.net/ecoinformas/boletin/pie.jpg' useMap=#Map border=0></DIV></DIV>"&chr(13)
		cuerpo = cuerpo & "<DIV>"&chr(13)
		cuerpo = cuerpo & "<DIV id=sombra_abajo></DIV>"&chr(13)
		cuerpo = cuerpo & "</DIV></DIV></DIV>"&chr(13)
		cuerpo = cuerpo & "</BODY></HTML>"

		if email<>"" then
			Set Mail = Server.CreateObject("Persits.MailSender")
			Mail.Host = "smtp.istas.net"
			Mail.From = "jdejong@istas.net"
			Mail.FromName = "ECOinformas" ' Opcional 
			Mail.Username = "say5151"
			Mail.Password = "***REMOVED**"
			Mail.AddAddress email
			Mail.Subject = Mail.EncodeHeader("Recordatorio de acceso a ECOinformas")
			Mail.Body = cuerpo
			Mail.IsHTML = True
			On Error Resume Next
			Mail.Send      ' ó Mail.SendToQueue
			If Err <> 0 Then
	      			Response.Write "Error en la cuenta " & email & ": " & Err.Description & "<br>" 
			End If 
		end if

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
<title>ECOinformas: Base de datos de sustancias tóxicas y peligrosas RISCTOX</title>
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
			<div id="encabezado_nuevo3">
			<table width="100%" cellpadding=0 border=0>
			<tr><td width="215" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="142" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="166" height="78" onclick="location.href='index.asp?idpagina=549'" style="cursor:hand">&nbsp;</td>
			    <td width="160" height="78" onclick="location.href='index.asp?idpagina=550'" style="cursor:hand">&nbsp;</td>
			    <td width="25"  height="78" align="center">
			    	<a href="mailto:salvira@istas.ccoo.es?subject=Contacto ECOinformas"><img src="imagenes/ico_contacto.gif" border="0" alt="Contacto"></a><br>
			    	<a href="busqueda.asp"><img src="imagenes/ico_busqueda.gif" border="0" alt="busqueda"></a><br>
			    	<a href="index.asp?idpagina=560"><img src="imagenes/ico_ayuda.gif" border="0" alt="ayuda"></a>
			    </td>
			</tr>
			</table>
			</div>
			<div id="menusup3">
			<table border="0" cellspacing="5" cellpadding="0">
            			<tr class="textmenusup">
<%              				sql = "SELECT titulo,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion LIKE 'AIC%' AND len(numeracion)=4 ORDER BY numeracion"
			   	   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	           	set objRecordset = OBJConnection.Execute(sql)
		   	   	           	do while not objRecordset.eof
              						response.write "<td class=textmenusup>"
							if mid(numeracion,1,4)=mid(objRecordset("numeracion"),1,4) then
								response.write lcase(objRecordset("titulo"))
              						else
              							response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&" style='text-decoration:none'>"&lcase(objRecordset("titulo"))&"</a>"
              						end if
              						response.write "</td><td class=textmenusup>|</td>"
							objrecordset.movenext
 						loop %>
              			</tr>
          		</table>
			</div>
			<div class="textsubmenu" id="submenusup3">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
            			<tr><td align="right">Visita virtual&nbsp;</td></tr>
          		</table>
			</div>
			
			<div id="texto">
			
				<div class="texto">
				
				<p class=titulo3>Fin de visita virtual</p>
				<div id="texto">
					<div class="texto">
				<br>&nbsp;
				<table width="90%" align="center">
				<tr><td class="texto">Gracias por registrarte como usuario de ECOInformas*. Con la clave y contraseña que te asignamos a continuación, tendrás acceso libre a toda la página web y podrás aprovechar los servicios y productos que te ofrecemos<br><br>
				CLAVE:&nbsp;<b><%=clave%></b><br>CONTRASEÑA:&nbsp;<b><%=contra%></b><br><br>Te hemos enviado un email con la clave y contraseña a la dirección <i><%=email%></i><br><br>
				<%=texto1%></td></tr>
				<tr><td class="texto"><b>Ahora ya puedes entrar en cualquier parte de esta web. La próxima vez que entres recuerda introducir tu clave y contraseña.</b></td></tr>
				</table>
			  <% if 1=0 then %>
			  <p align="center"><input type="button" value="DARSE DE ALTA EN ALGÚN CURSO" class="boton" onclick="location.href='formulario_identificado.asp'"></p>
			  <% end if %>
				
				<center><object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" width=530 height=100 id="anima0" align="middle"><param name="allowScriptAccess" value="sameDomain" /><param name="movie" value="http://www.istas.net/recursos/ANI/ISTAS_01076.swf" /><param name="quality" value="high" /><param name="wmode" value="transparent" /><param name="bgcolor" value="#ffffff" /><embed src="http://www.istas.net/recursos/ANI/ISTAS_01076.swf" quality="high" wmode="transparent" bgcolor="#ffffff" width=530 height=100 name="" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" /></object></center>
						<p>&nbsp;</p>
					</div>
				</div>



				</div>
				</div>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
			</div>

     			<map name="Map3" id="Map3">
            		<area shape="rect" coords="300,18,392,80" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="393,18,539,80" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,18,694,80" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<img src="imagenes/pie3.jpg" width="708" border="0" usemap="#Map3">
      			
    			
    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>
</body>
</html>
