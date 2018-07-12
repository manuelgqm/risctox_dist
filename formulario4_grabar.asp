<!--#include file="EliminaInyeccionSQL.asp"-->
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"



	dim campo(65,4)
	campo(1,0)="nombre"        
	campo(2,0)="apellidos"        
	campo(3,0)="fec_nac"    
	campo(4,0)="sexo" 
	'campo(5,0)="seg_social"       
	campo(6,0)="minusvalia"       
	campo(7,0)="inmigrante"       
	campo(8,0)="cualificacion"   
	'campo(9,0)="dni"   
	campo(10,0)="cond_laboral"   
	campo(11,0)="tam_empresa" 
	'campo(12,0)="puesto" 
	'campo(13,0)="contrato" 
	'campo(14,0)="estudios" 
	campo(15,0)="direccion" 
	campo(16,0)="localidad" 
	campo(17,0)="provincia" 
	campo(18,0)="cp" 
	campo(19,0)="telefono" 
	'campo(20,0)="movil"
	'campo(21,0)="fax"
	campo(22,0)="email"  
	campo(23,0)="empresa"     
	'campo(24,0)="cif"  
	'campo(25,0)="razon_social"  
	'campo(26,0)="sector"     
	'campo(27,0)="emp_direccion"  
	'campo(28,0)="emp_localidad"     
	campo(29,0)="emp_provincia"  
	'campo(30,0)="emp_cp"     
	'campo(31,0)="emp_telefono"  
	'campo(32,0)="emp_movil"     
	'campo(33,0)="emp_fax"  
	'campo(34,0)="emp_email"     
	'campo(35,0)="emp_web"  
	campo(36,0)="recibir_info_ecoinformas"     
	campo(37,0)="recibir_info_istas"  
	campo(38,0)="observaciones"     
	'campo(39,0)="FP02"  
	'campo(40,0)="FP03"     
	'campo(41,0)="FP04"  
	'campo(42,0)="FP01"    
	'campo(43,0)="FDT01" 
	'campo(44,0)="SJ01"    
	campo(39,0)="FDT01"  
	campo(40,0)="FDT02"     
	campo(41,0)="FDT03"  
	campo(42,0)="FDT04"    
	campo(43,0)="FDT05" 
	campo(44,0)="FDT06"    
	
	'campo(45,0)="SJ02" 
	campo(46,0)="direccion_materiales"     
	'campo(47,0)="FolGen"   
	'campo(48,0)="FolObs"   
	campo(49,0)="SET04"   
	campo(50,0)="EGP01"   
	campo(51,0)="EGP02"   
	campo(52,0)="EGP03"   
	campo(53,0)="EGP04"   
	campo(54,0)="EGP05"   
	campo(55,0)="EGP06"   
	campo(56,0)="EGP07"   
	campo(57,0)="AE01"   
	campo(58,0)="AE02"   
	campo(59,0)="AE03"   
	campo(60,0)="AE04"   
	campo(61,0)="AE05"   
	campo(62,0)="AE06"   
	campo(63,0)="SEP01"
	campo(64,0)="SEP02"
	campo(65,0)="SEP03"
	
	
	'for i=1 to 62
	'response.write "if (c"&i&"=='') { falta = falta+'"&campo(i,0)&"'+'\n'; }<br>"
	'next
	
	campo(4,1)="1"   
	campo(6,1)="1"   
	campo(7,1)="1"   
	campo(8,1)="1"   
	campo(10,1)="1"   
	campo(11,1)="1"   
	campo(12,1)="1"   
	campo(13,1)="1"   
	campo(14,1)="1"   
	campo(17,1)="1"  
	campo(26,1)="1"  
	campo(29,1)="1"  
	campo(36,1)="1"  
	campo(37,1)="1"  
	campo(46,1)="1" 
	
	campo(39,1)="2" 
	campo(40,1)="2" 
	campo(41,1)="2" 
	campo(42,1)="2" 
	campo(43,1)="2" 
	campo(44,1)="2" 
	campo(45,1)="2" 
	campo(47,1)="2" 
	campo(48,1)="2" 
	campo(49,1)="2" 
	campo(50,1)="2" 
	campo(51,1)="2" 
	campo(52,1)="2" 
	campo(53,1)="2" 
	campo(54,1)="2" 
	campo(55,1)="2" 
	campo(56,1)="2" 
	campo(57,1)="2" 
	campo(58,1)="2" 
	campo(59,1)="2" 
	campo(60,1)="2" 
	campo(61,1)="2" 
	campo(62,1)="2" 
	campo(63,1)="2"
	campo(64,1)="2"
	campo(65,1)="2"

	campo(1,3)="Nombre"        
	campo(2,3)="Apellidos"        
	campo(3,3)="Fecha nacimiento"    
	campo(4,3)="Sexo" 
	campo(5,3)="Seguridad social"       
	campo(6,3)="Minusvalía"       
	campo(7,3)="Inmigrante"       
	campo(8,3)="Baja cualificación"   
	campo(9,3)="DNI/NIE"   
	campo(10,3)="Condición laboral"   
	campo(11,3)="Tamaño empresa" 
	campo(12,3)="Puesto" 
	campo(13,3)="Contrato" 
	campo(14,3)="Estudios" 
	campo(15,3)="Dirección" 
	campo(16,3)="Localidad" 
	campo(17,3)="Provincia" 
	campo(18,3)="CP" 
	campo(19,3)="Teléfono" 
	campo(20,3)="Movil"
	campo(21,3)="Fax"
	campo(22,3)="Email"  
	campo(23,3)="Empresa"     
	campo(24,3)="CIF"  
	campo(25,3)="Razón social"  
	campo(26,3)="Sector"     
	campo(27,3)="Dirección"  
	campo(28,3)="Localidad"     
	campo(29,3)="Provincia"  
	campo(30,3)="CP"     
	campo(31,3)="Teléfono"  
	campo(32,3)="Movil"     
	campo(33,3)="Fax"  
	campo(34,3)="Email"     
	campo(35,3)="Web"  
	campo(36,3)="Recibir información de ECOinformas"     
	campo(37,3)="Recibir información de ISTAS"  
	campo(38,3)="Observaciones"     
	'campo(39,3)="FP02"  
	'campo(40,3)="FP03"     
	'campo(41,3)="FP04"  
	'campo(42,3)="FP01"    
	'campo(43,3)="FDT01" 
	'campo(44,3)="SJ01" 
	campo(39,3)="FDT01"  
	campo(40,3)="FDT02"     
	campo(41,3)="FDT03"  
	campo(42,3)="FDT04"    
	campo(43,3)="FDT05" 
	campo(44,3)="FDT06"
	   
	campo(45,3)="SJ02" 
	campo(46,3)="Dirección materiales"     
	campo(47,3)="FolGen"   
	campo(48,3)="FolObs"   
	campo(49,3)="SET04"   
	campo(50,3)="EGP01"   
	campo(51,3)="EGP02"   
	campo(52,3)="EGP03"   
	campo(53,3)="EGP04"   
	campo(54,3)="EGP05"   
	campo(55,3)="EGP06"   
	campo(56,3)="EGP07"   
	campo(57,3)="AE01"   
	campo(58,3)="AE02"   
	campo(59,3)="AE03"   
	campo(60,3)="AE04"   
	campo(61,3)="AE05"   
	campo(62,3)="AE06"   
	campo(63,3)="SEP01"
	campo(64,3)="SEP02"
	campo(65,3)="SEP03"

	'-- Campos útiles (18-enero-2006)
	campo(1,4)="1"
	campo(2,4)="1"
	campo(3,4)="1"
	campo(4,4)="1"
	campo(5,4)="0"
	campo(6,4)="1"
	campo(7,4)="1"
	campo(8,4)="1"
	campo(9,4)="0"
	campo(10,4)="1"
	campo(11,4)="1"
	campo(12,4)="0"
	campo(13,4)="0"
	campo(14,4)="0"
	campo(15,4)="1"
	campo(16,4)="1"
	campo(17,4)="1"
	campo(18,4)="1"
	campo(19,4)="1"
	campo(20,4)="0"
	campo(21,4)="0"
	campo(22,4)="1"
	campo(23,4)="1"
	campo(24,4)="0"
	campo(25,4)="0"
	campo(26,4)="0"
	campo(27,4)="0"
	campo(28,4)="0"
	campo(29,4)="1"
	campo(30,4)="0"
	campo(31,4)="0"
	campo(32,4)="0"
	campo(33,4)="0"
	campo(34,4)="0"
	campo(35,4)="0"
	campo(36,4)="1"     
	campo(37,4)="1"
	campo(38,4)="1"
	campo(39,4)="1"
	campo(40,4)="1"
	campo(41,4)="1"
	campo(42,4)="1"
	campo(43,4)="1"
	campo(44,4)="1"
	campo(45,4)="0"
	campo(46,4)="0"
	campo(47,4)="0"
	campo(48,4)="0"
	campo(49,4)="0"
	campo(50,4)="0"
	campo(51,4)="0"
	campo(52,4)="0"
	campo(53,4)="0"
	campo(54,4)="0"
	campo(55,4)="0"
	campo(56,4)="0"
	campo(57,4)="0"
	campo(58,4)="0"
	campo(59,4)="0"
	campo(60,4)="0"
	campo(61,4)="0"
	campo(62,4)="0"
	campo(63,4)="0"
	campo(64,4)="0"
	campo(65,4)="0"
	
for i=1 to 65
  if campo(i,4)="1" then
	if campo(i,1)="1" then
		campo(i,2) = valor(EliminaInyeccionSQL(request(campo(i,0))))
	else
		if campo(i,1)="2" then
			if EliminaInyeccionSQL(request(campo(i,0)))="1" then
				campo(i,2)="sí"
			else
				campo(i,2)="no"
			end if
		else			
			campo(i,2) = unquote(EliminaInyeccionSQL(request(campo(i,0))))
		end if
	end if
  end if		
next

orden = "INSERT ECOINFORMAS_GENTE ("
for i=1 to 65
	if campo(i,4)="1" then orden = orden & campo(i,0) & ","
next
orden = orden & "fec_hor,ip,clave,contra,fec_hor_mod,usu_mod,confirmado_web,confirmado_cursos,confirmado_materiales) VALUES ('"
for i=1 to 65
	if campo(i,4)="1" then orden = orden & unquote(EliminaInyeccionSQL(request(campo(i,0)))) & "','"
next
orden = orden & now() & "','" & Request.ServerVariables("REMOTE_ADDR") & "','','','" & now() & "',0,0,0,0);"
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
Set objRecordset = OBJConnection.Execute(orden)

orden = "SELECT max(idgente) as ultimo FROM ECOINFORMAS_GENTE"
Set objRecordset = OBJConnection.Execute(orden)
ultimo = objRecordset("ultimo")
session("id_ecogente") = cstr(ultimo)

clave = "ECO"&(ultimo)
apellidos = split(EliminaInyeccionSQL(request(campo(2,0)))," ")
contra = mid(trim(ucase(apellidos(0))),1,15)

orden = "UPDATE ECOINFORMAS_GENTE set clave='"&clave&"',contra='"&contra&"' WHERE idgente="&ultimo
Set objRecordset = OBJConnection.Execute(orden)

email = EliminaInyeccionSQL(request("email"))

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
		'cuerpo = cuerpo&"<P>&nbsp;</P>"&chr(10)& chr(13)
		
		cuerpo = cuerpo & "<p class=texto align=left>Estimada señora, estimado señor:</p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left><b>Gracias por registrarse como usuario de ECOInformas*.</b></p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Con la clave y contraseña que le asignamos a continuación, tendrá acceso libre a toda la página web y podrá aprovechar los servicios y productos gratuitos que le ofrecemos.</p>" & chr(13)
		cuerpo = cuerpo & "<p class='texto' align='center'>Clave: <b>"&clave&"</b><br>Contraseña: <b>"&contra&"</b></p>" & chr(13)
		cuerpo = cuerpo & "<p class=texto align=left><b>¿Sabe cómo prevenir el riesgo químico en su trabajo?</b></p>"&chr(13)
		cuerpo = cuerpo & "<p class=texto align=left>Entre en <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=575'><b>RISCTOX</b></a> con su clave y contraseña e infórmese sobre las sustancias químicas que hay en su centro de trabajo y sobre las <a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=576'><b>Alternativas</b></a> de sustiución para proteger mejor su salud y el medio ambiente.</p>"&chr(13)
		'cuerpo = cuerpo & "<p class=texto align=left><a href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=521'><b>¿Conoce lo que usa?</b></a> Con esta guía onlina le ayudamos a conseguir información sobre las sustancias químicas presentes en su centro de trabajo.</p>" & chr(10)& chr(13)
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
		'cuerpo = cuerpo & "<p class=texto align=left>" & texto2 & "</p>" & chr(10)& chr(13)
		
		cuerpo = cuerpo & "<P>&nbsp;</P>"&chr(13)

		cuerpo = cuerpo & "<p align=center><map name='mapa_poste'><area href='http://www.ecoinformas.com' shape='rect' coords='41, 21, 195, 66'><area href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=575' shape='rect' coords='87, 79, 233, 119'><area href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=576' shape='rect' coords='19, 130, 142, 168'><area href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=961' shape='rect' coords='88, 174, 233, 216'><area href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=710' shape='rect' coords='16, 227, 149, 263'><area href='http://www.istas.net/ecoinformas/web/index.asp?idpagina=566' shape='rect' coords='99, 268, 229, 311'></map>"&chr(13)
		cuerpo = cuerpo & "<img border='0' src='http://www.istas.net/ecoinformas/web/imagenes/eco_poste.gif' usemap='#mapa_poste' width='242px' height='362px'></p>"
		cuerpo = cuerpo & "<P>&nbsp;</P>"&chr(13)
		
		cuerpo = cuerpo & "<MAP name='Map'>"& chr(10)& chr(13)
		cuerpo = cuerpo & "<AREA shape='RECT' target='_blank' alt='Fundación Biodiversidad' coords='267,35,339,84' href='http://www.fundacion-biodiversidad.es'>"
		cuerpo = cuerpo & "<AREA shape='RECT' target='_blank' alt='Instituto Sindical de Trabajo, Ambiente y Salud' coords='342,34,471,84' href='http://www.istas.ccoo.es'>"
		cuerpo = cuerpo & "<AREA shape='RECT' target='_blank' alt='Fondo Social Europeo' coords='472,35,599,82' href='http://www.mtas.es/UAFSE/default.htm'>"
		cuerpo = cuerpo & "</MAP>"&chr(13)
		cuerpo = cuerpo & "<DIV><IMG src='http://www.istas.net/ecoinformas/boletin/pie.jpg' useMap='#Map' border='0'></DIV></DIV>"&chr(13)
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


FUNCTION valor(v)
	if cstr(v)<>"0" and cstr(v)<>"" then
		orden2 = "SELECT desc1 FROM ECOINFORMAS_VALORES WHERE valor="&v
		Set dSQL2 = Server.CreateObject ("ADODB.Recordset")
		dSQL2.Open orden2,objConnection,adOpenKeyset
		valor = dSQL2("desc1")
	else
		valor = "sin especificar"
	end if
END FUNCTION

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: grabado formulario</title>
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
			<div id="encabezado_nuevo1">
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
			<div id="menusup1">
			<table border="0" cellspacing="5" cellpadding="0">
            			<tr class="textmenusup"><td class=textmenusup>Formulario</td>
          		</table>
			</div>
			
			<div class="textsubmenu" id="submenusup1">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
            			<tr>
              				<td width="100%" valign="top">Est&aacute;s en: Formulario para solicitar acceso a la web y solicitar materiales</td>
            			</tr>
          		</table>
			</div>
			
			<div id="texto">
				<div class="texto">
				
				<br>&nbsp;
				<table width="90%" align="center" clas=tabla>
				<tr><td class="texto" align="center">Gracias por remitir tus datos. Tus claves de acceso son:
				<p class="texto" width="50%">CLAVE:&nbsp;<b><%=clave%></b></p>
				<p class="texto" width="50%">CONTRASEÑA:&nbsp;<b><%=contra%></b></p>
				Acabamos de enviarte un correo a la dirección <%=email%> con tus claves de acceso para que puedas volver a entrar en todas las herramientas de ECOinformas.<br>
				</td></tr></table>
				
				<p class=subtitulo>Solicitud de cursos ECOinformas</p>
				<p>Si perteneces a alguno de los <a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=557" target="_blank">colectivos a los que van dirigidos los cursos on-line</a> puedes apuntarte y realizarlos de forma gratuita:</p>
				<p align="center"><input type="button" value="SOLICITUD CURSOS ON-LINE" class="boton" onclick="location.href='formulario_identificado.asp'"></p>
				<br><br>
				
				<p class=subtitulo>Visita virtual</p>
				<p>Además te ofrecemos una visita virtual al espacio web, para que descubras las herramientas útiles que contiene. Son sólo 8 pasos. Pulsa el botón <img src="imagenes/avancevisita.gif" align="absmiddle">&nbsp;para avanzar en la visita o usa los números de la izquierda.</p>
				<p align="center"><input type="button" value="INICIAR LA VISITA VIRTUAL" class="boton" onclick="location.href='visita.asp'"></p>

				<br>&nbsp;
								
				<table class=tabla width="90%" align="center">
				<tr><td class="subtitulo" colspan="2">Datos personales</td></tr>
					<% for i=1 to 65 %>
					<% if campo(i,4)="1" then %>
					<tr><td class="celda" align="right"><%=campo(i,3)%>:&nbsp;</td><td class="celda"><%=campo(i,2)%></td></tr>
					<% end if %>

					<% if i=36 then %>
					<tr><td class="celda" colspan="2">&nbsp;</td></tr>
					<tr><td class="subtitulo" colspan="2">Otros datos/opciones</td></tr>
					<% end if %>

					<% if i=46 then %>
					<tr><td class="celda" colspan="2">&nbsp;</td></tr>
					<tr><td class="subtitulo" colspan="2">Solicitud de envío de materiales</td></tr>
					<% end if %>
				<% next %>
				</table>
				
				<br>&nbsp;
				
				<table width="90%" align="center">
				<tr><td class=texto><%=texto2%></td></tr>
				<tr><td class="texto">&nbsp;</td></tr>
				</table>
				
				</div>
				<p align="center"><input type="button" value="IMPRIMIR" class="boton" onClick="print()"></p>
				<p>&nbsp;</p>
			</div>

			<map name="Map1" id="Map1">
            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="400,38,546,104" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="547,38,701,104" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<img src="imagenes/pie1.jpg" width="708" border="0" usemap="#Map1">

    			</div>
    		</div>
		<div id="sombra_abajo"></div>
	</div>
</body>
</html>