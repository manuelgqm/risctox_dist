<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

   if session("id_ecogente")="" then response.redirect "formulario2.asp"
	
   seg_social = request("seg_social")
   dni = request("dni")
   puesto = request("puesto")
   contrato = request("contrato")
   estudios = request("estudios")
   direccion = request("direccion")
   localidad = request("localidad")
   provincia = request("provincia")
   cp = request("cp")
   movil = request("movil")
   fax = request("fax")
   empresa = request("empresa")
   cif = request("cif")
   sector = request("sector")
   emp_direccion = request("emp_direccion")
   emp_localidad = request("emp_localidad")
   emp_cp = request("emp_cp")
   emp_telefono = request("emp_telefono")
   emp_fax = request("emp_fax")
   emp_movil = request("emp_movil")
   emp_web = request("emp_web")
   emp_email = request("emp_email")
   relacion_ma = request("relacion_ma")
   rlt = request("rlt")
   emp_tipo = request("emp_tipo")
   emp_facturacion = request("emp_facturacion")
   emp_provincia = request("emp_provincia")
	
   seg_social = unquote(seg_social)
   dni = unquote(dni)
   puesto = unquote(puesto)
   contrato = unquote(contrato)
   estudios = unquote(estudios)
   direccion = unquote(direccion)
   localidad = unquote(localidad)
   provincia = unquote(provincia)
   cp = unquote(cp)
   movil = unquote(movil)
   fax = unquote(fax)
   empresa = unquote(empresa)
   cif = unquote(cif)
   sector = unquote(sector)
   emp_direccion = unquote(emp_direccion)
   emp_localidad = unquote(emp_localidad)
   emp_cp = unquote(emp_cp)
   emp_telefono = unquote(emp_telefono)
   emp_fax = unquote(emp_fax)
   emp_movil = unquote(emp_movil)
   emp_web = unquote(emp_web)
   emp_email = unquote(emp_email)
   relacion_ma = unquote(relacion_ma)
   rlt = unquote(rlt)
   emp_tipo = unquote(emp_tipo)
   emp_facturacion = unquote(emp_facturacion)
   emp_provincia = unquote(emp_provincia)
	
FDT01_2007 = request("FDT01_2007")
if FDT01_2007<>"" then FDT01_2007=1 else FDT01_2007=0
FDT02_2007 = request("FDT02_2007")
if FDT02_2007<>"" then FDT02_2007=1 else FDT02_2007=0
FDT03_2007 = request("FDT03_2007")
if FDT03_2007<>"" then FDT03_2007=1 else FDT03_2007=0
FDT04_2007 = request("FDT04_2007")
if FDT04_2007<>"" then FDT04_2007=1 else FDT04_2007=0
FDT05_2007 = request("FDT05_2007")
if FDT05_2007<>"" then FDT05_2007=1 else FDT05_2007=0
FDT06_2007 = request("FDT06_2007")
if FDT06_2007<>"" then FDT06_2007=1 else FDT06_2007=0
	

orden = "UPDATE ECOINFORMAS_GENTE SET FDT01_2007='"&FDT01_2007&"',FDT02_2007='"&FDT02_2007&"',FDT03_2007='"&FDT03_2007&"',FDT04_2007='"&FDT04_2007&"',FDT05_2007='"&FDT05_2007&"',FDT06_2007='"&FDT06_2007&"',"
orden = orden & "seg_social='"&seg_social&"',"
orden = orden & "dni='"&dni&"',"
orden = orden & "puesto='"&puesto&"',"
orden = orden & "contrato='"&contrato&"',"
orden = orden & "estudios='"&estudios&"',"
orden = orden & "direccion='"&direccion&"',"
orden = orden & "localidad='"&localidad&"',"
orden = orden & "provincia='"&provincia&"',"
orden = orden & "cp='"&cp&"',"
orden = orden & "movil='"&movil&"',"
orden = orden & "fax='"&fax&"',"
orden = orden & "empresa='"&empresa&"',"
orden = orden & "cif='"&cif&"',"
orden = orden & "sector='"&sector&"',"
orden = orden & "emp_direccion='"&emp_direccion&"',"
orden = orden & "emp_localidad='"&emp_localidad&"',"
orden = orden & "emp_cp='"&emp_cp&"',"
orden = orden & "emp_telefono='"&emp_telefono&"',"
orden = orden & "emp_fax='"&emp_fax&"',"
orden = orden & "emp_movil='"&emp_movil&"',"
orden = orden & "emp_web='"&emp_web&"',"
orden = orden & "emp_email='"&emp_email&"',"
orden = orden & "relacion_ma='"&relacion_ma&"',"
orden = orden & "rlt='"&rlt&"',"
orden = orden & "emp_tipo='"&emp_tipo&"',"
orden = orden & "emp_facturacion='"&emp_facturacion&"',"
orden = orden & "emp_provincia='"&emp_provincia&"',"
orden = orden & "fec_hor_mod='"&now()&"',"
orden = orden & "usu_mod='"&session("id_ecogente")&"' "
orden = orden & " WHERE idgente="&session("id_ecogente")
Set objRecordset = OBJConnection.Execute(orden)
'response.write orden



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
			
			<% if session("id_ecogente")<>"" then %>
			<div class="textsubmenu" id="submenusup<% response.write (seccion) %>">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
<%            			sql = "SELECT nombre,apellidos,sexo FROM ECOINFORMAS_GENTE WHERE idgente="&session("id_ecogente")
			   	   	      set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	        set objRecordset = OBJConnection.Execute(sql)
		   	   	        usuario = objRecordset("nombre")&" "&objRecordset("apellidos")
		   	   	        usuario_sexo = "o"
		   	   	        if objRecordset("sexo")=75 then usuario_sexo = "a"
%>
            			<tr><td align="right">Usuari<%=usuario_sexo%> identificad<%=usuario_sexo%>:&nbsp;<%=usuario%>&nbsp;</td></tr>
          		</table>
			</div>
       			<% end if %>
			
			<div id="texto">
				<div class="texto">
				
				<br>&nbsp;
				
				<table width="90%" align="center">
				<tr><td class="texto" colspan="5">Tus datos se han grabado correctamente. En breve nos pondremos en contacto contigo para informarte sobre los cursos que has solicitado.<br>&nbsp;<br>&nbsp;</td></tr>
				<% if FDT01_2007<>"0" or FDT02_2007<>"0" or FDT03_2007<>"0" or FDT04_2007<>"0" or FDT05_2007<>"0" or FDT06_2007<>"0" then %>
				<tr><td>&nbsp;</td><td><strong>Código</strong></td><td><strong>Acción formativa</strong></td></tr>
				<% if FDT01_2007<>"0" then %>
				<tr><td class="celda">&nbsp;</td><td class="celda">FDT01</td><td class="celda">Curso on-line sobre "Introducción a los Sistemas de Gestión Medioambiental"</td></tr>
				<% end if %>
				<% if FDT02_2007<>"0" then %>
				<tr><td class="celda">&nbsp;</td><td class="celda">FDT02</td><td class="celda">Curso on-line sobre "Introducción a los Sistemas de Gestión Medioambiental"</td></tr>
				<% end if %>
				<% if FDT03_2007<>"0" then %>
				<tr><td class="celda">&nbsp;</td><td class="celda">FDT03</td><td class="celda">Curso on-line Básico de Riesgo Químico</td></tr>
				<% end if %>
				<% if FDT04_2007<>"0" then %>
				<tr><td class="celda">&nbsp;</td><td class="celda">FDT04</td><td class="celda">Curso on-line Básico de Riesgo Químico</td></tr>
				<% end if %>
				<% if FDT05_2007<>"0" then %>
				<tr><td class="celda">&nbsp;</td><td class="celda">FDT05</td><td class="celda">Curso on-line sobre "Energías Renovables, Medio Ambiente y Empleo"</td></tr>
				<% end if %>
				<% if FDT06_2007<>"0" then %>
				<tr><td class="celda">&nbsp;</td><td class="celda">FDT06</td><td class="celda">Curso on-line sobre "Energías Renovables, Medio Ambiente y Empleo"</td></tr>
				<% end if %>
				<% end if %>
				
				</table>
			  <br>&nbsp;
				<br>&nbsp;
				
			  <%if 1=0 then %><p align="center">Si quieres recibir todos tus datos en tu correo electrónico pulsa el botón <input type="button" class="boton" value="COMPROBAR MIS DATOS" onclick="location.href='pedir_datos.asp'"></p><% end if %>
				
				<center><object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" width=530 height=100 id="anima0" align="middle"><param name="allowScriptAccess" value="sameDomain" /><param name="movie" value="http://www.istas.net/recursos/ANI/ISTAS_01076.swf" /><param name="quality" value="high" /><param name="wmode" value="transparent" /><param name="bgcolor" value="#ffffff" /><embed src="http://www.istas.net/recursos/ANI/ISTAS_01076.swf" quality="high" wmode="transparent" bgcolor="#ffffff" width=530 height=100 name="" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" /></object></center>
				
				<br>&nbsp;
								
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