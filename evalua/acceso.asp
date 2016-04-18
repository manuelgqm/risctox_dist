<!--#include file="../dn_conexion.asp"-->

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

<link rel="stylesheet" type="text/css" href="estructura.css">
</head>
<body>
<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
		<!--#include file="../dn_cabecera.asp"-->
		<div id="texto">
			
			<div class="texto">
				<table width="100%" border="0">
                <tr>
                <td></td>
                <td align='right'><input type="button" name="volver" class="boton" value="Volver a la portada de evalúa lo que usas" onclick="window.location='./index.asp';"></td>
                </tr>
                </table>
				<p class="titulo3">P&aacute;gina de identificación</p>
			
			
				Est&aacute;s en: identificaci&oacute;n para acceso a zonas restringidas
            
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
                    
                    <% if request("recordar")=1 then %>
					<div class="tabla" style='width:90%;height:50px'>
					<p class="texto">Acabamos de enviarte un correo electrónico a tu dirección con la clave y contraseña para que puedas acceder libremente a todo el espacio web de Riesgo Químico.</p>
					</div>
                    <% end if %>
					
					<div id="identifica">
					<form name="form1" id="form1" method="post" action="identifica.asp">
					<input type="hidden" name="idenlace" value="<%=request("idenlace")%>">
					<input type="hidden" name="idpagina" value="<%=request("idpagina")%>">
					<input type="hidden" name="pagina_esp" value="<%=request("pagina_esp")%>">
					<table border="0" cellspacing="2" cellpadding="2" align="center" bgcolor="#52B1DB">
				  		<tr bgcolor="#097DB0">
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
    			</div>	

		</div>
        </div>
        
			<img src="../imagenes/pie_risctox.gif" width="708" border="0">
			
    		
    	</div>
	<div id="sombra_abajo"></div>
</div>
</body>
</html>

<%
cerrarconexion
%>

