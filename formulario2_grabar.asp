<!--#include file="EliminaInyeccionSQL.asp"-->
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

	'usuario = request.cookies("webistas")

	dim campo(62,3)
	campo(1,0)="nombre"        
	campo(2,0)="apellidos"        
	campo(3,0)="fec_nac"    
	campo(4,0)="sexo" 
	campo(5,0)="seg_social"       
	campo(6,0)="minusvalia"       
	campo(7,0)="inmigrante"       
	campo(8,0)="cualificacion"   
	campo(9,0)="dni"   
	campo(10,0)="cond_laboral"   
	campo(11,0)="tam_empresa" 
	campo(12,0)="puesto" 
	campo(13,0)="contrato" 
	campo(14,0)="estudios" 
	campo(15,0)="direccion" 
	campo(16,0)="localidad" 
	campo(17,0)="provincia" 
	campo(18,0)="cp" 
	campo(19,0)="telefono" 
	campo(20,0)="movil"
	campo(21,0)="fax"
	campo(22,0)="email"  
	campo(23,0)="empresa"     
	campo(24,0)="cif"  
	campo(25,0)="razon_social"  
	campo(26,0)="sector"     
	campo(27,0)="emp_direccion"  
	campo(28,0)="emp_localidad"     
	campo(29,0)="emp_provincia"  
	campo(30,0)="emp_cp"     
	campo(31,0)="emp_telefono"  
	campo(32,0)="emp_movil"     
	campo(33,0)="emp_fax"  
	campo(34,0)="emp_email"     
	campo(35,0)="emp_web"  
	campo(36,0)="recibir_info_ecoinformas"     
	campo(37,0)="recibir_info_istas"  
	campo(38,0)="observaciones"     
	campo(39,0)="FP02"  
	campo(40,0)="FP03"     
	campo(41,0)="FP04"  
	campo(42,0)="FP01"    
	campo(43,0)="FDT01" 
	campo(44,0)="SJ01"    
	campo(45,0)="SJ02" 
	campo(46,0)="direccion_materiales"     
	campo(47,0)="FolGen"   
	campo(48,0)="FolObs"   
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
	campo(39,3)="FP02"  
	campo(40,3)="FP03"     
	campo(41,3)="FP04"  
	campo(42,3)="FP01"    
	campo(43,3)="FDT01" 
	campo(44,3)="SJ01"    
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
	
for i=1 to 62
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
next

orden = "INSERT ECOINFORMAS_GENTE ("
for i=1 to 62
	orden = orden & campo(i,0) & ","
next
orden = orden & "fec_hor,ip,clave,contra,confirmado_web,confirmado_cursos,confirmado_materiales) VALUES ('"
for i=1 to 62
	orden = orden & unquote(rEliminaInyeccionSQL(equest(campo(i,0)))) & "','"
next
orden = orden & now() & "','" & Request.ServerVariables("REMOTE_ADDR") & "','','',0,0,0);"
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
Set objRecordset = OBJConnection.Execute(orden)



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
			
			<div class="textsubmenu" id="submenusup">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
            			<tr>
              				<td width="100%" valign="top">Est&aacute;s en: Formulario para solicitar acceso a la web, matricularte en los cursos y solicitar materiales</td>
            			</tr>
          		</table>
			</div>
			
			<div id="texto">
				<div class="texto">
				
				<br>&nbsp;
				<table width="90%" align="center">
				<tr><td class="texto">Tus datos han sido mandados correctamente. El proyecto ECOinformas está concebido especialmente para los trabajadores de PYMEs españolas y profesionales cuya actividad incida de alguna manera en la protección del medio ambiente. Para poder recibir los servicios solicitados es necesario que verifiquemos si perteneces a estos colectivos, que seas trabajador o trabajadora, director o directora, técnico, responsable medioambiental o representante de los trabajadores y trabajadoras, si eres autónomo o perteneces a colectivos desfavorecidos. Una vez verificada esta información recibirás un correo electrónico con las claves y contraseñas correspondientes para que puedas acceder al servicio. <b>Este proceso puede tardar un par de días ya que se realiza personalmente</b>.</td></tr>
				</table>
				
				<br>&nbsp;
								
				<table class=tabla width="90%" align="center">
				<tr><td class="subtitulo" colspan="2">Datos personales</td></tr>
					<% for i=1 to 62 %>
					<tr><td class="celda" align="right"><%=campo(i,3)%>:&nbsp;</td><td class="celda"><%=campo(i,2)%></td></tr>
					<% if i=22 then %>
					<tr><td class="celda" colspan="2">&nbsp;</td></tr>
					<tr><td class="subtitulo" colspan="2">Datos de la empresa en que trabajas</td></tr>
					<% end if %>
					<% if i=35 then %>
					<tr><td class="celda" colspan="2">&nbsp;</td></tr>
					<tr><td class="subtitulo" colspan="2">Otros datos/opciones</td></tr>
					<% end if %>
					<% if i=38 then %>
					<tr><td class="celda" colspan="2">&nbsp;</td></tr>
					<tr><td class="subtitulo" colspan="2">Solicitudes de inscripción en cursos</td></tr>
					<% end if %>
					<% if i=45 then %>
					<tr><td class="celda" colspan="2">&nbsp;</td></tr>
					<tr><td class="subtitulo" colspan="2">Solicitudes de envío de materiales</td></tr>
					<% end if %>
				<% next %>
				</table>
				
				<br>&nbsp;
				
				<table width="90%" align="center">
				<tr><td class="texto">Si has solicitado materiales o inscripción en algún curso deberás imprimir esta hoja, firmarla y enviarla por correo ordinario a la dirección de ISTAS con la referencia ECOinformas:<br>
				C/ General Cabrera, 21 E-28020 Madrid.</td></tr>	
				<tr><td class="texto">&nbsp;</td></tr>
				<tr><td class="texto"><b>Firma:</b>&nbsp;</td></tr>
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