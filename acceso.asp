<!--#include file="EliminaInyeccionSQL.asp"-->
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	'OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	'OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=217.13.81.22; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
	OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
	OBJConnection.Open

	'if session("risctox_en_webistas")="si" then session("id_ecogente") = 12471
	if 1=1 then session("id_ecogente") = 12471

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: identificación</title>
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
            			<tr class="textmenusup"><td class=textmenusup>P&aacute;gina de identificación</td>
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
		    			<% if request("error")=1 then %>
					<div class="tabla">
					<p class="subtitulo">Tu clave y contraseña no son válidas.<br>
					Vuelve a intentarlo o escribe a <a href="mailto:datospersonales@istas.net">datospersonales@istas.net</a> para comunicar el problema.</p>
					</div>
					<% end if %>
					
					<div id="identifica">
					<form name="form1" id="form1" method="post" action="identifica.asp">
					<input type="hidden" name="idenlace" value="<%=EliminaInyeccionSQL(request("idenlace"))%>">
					<input type="hidden" name="idpagina" value="<%=EliminaInyeccionSQL(request("idpagina"))%>">
					<input type="hidden" name="pagina_esp" value="<%=EliminaInyeccionSQL(request("pagina_esp"))%>">
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
					<form name="form_recordar" action="recordar.asp" method="POST">
					<p class="textoc">Si no recuerdas tu clave y contraseña, escribe aquí tu email y si coincide con el que te diste de alta,<br>te reenviaremos automáticamente el correo con tus datos de acceso.</p>
					<p class="textoc">Tu e-mail:&nbsp;<input type="text" class="campo" size="50" maxlenght="200" name="tu_email">&nbsp;<input class="boton" type="submit" name="Submit" value="Enviar" /></p>
				</form>
		    			<p>&nbsp;</p>
							
					<table width="80%" border="0" cellspacing="2" cellpadding="2" align="center" class="tabla">
					<tr><td class="textoc">Si todav&iacute;a no has solicitado tu clave para acceder a la web completa puedes rellenar el siguiente formulario. </td></tr>
					<tr><td class="textoc"><input type="button" class="boton" value="solicitar acceso libre" onclick="location.href='formulario2.asp'"/></td></tr></table>
	
		    			<p>&nbsp;</p>
		    			<p>&nbsp;</p>
		
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
