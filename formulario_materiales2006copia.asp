<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

	if session("id_ecogente")="" then response.redirect "acceso.asp"

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


FUNCTION valores (vgrupo,vname,tipo)

orden2 = "SELECT * FROM ECOINFORMAS_VALORES WHERE grupo='"&vgrupo&"' ORDER BY valor,desc1"
Set dSQL2 = Server.CreateObject ("ADODB.Recordset")
dSQL2.Open orden2,objConnection,adOpenKeyset
if tipo="1" then 				'--- tipo 1 = desplegable
%>	<select name=<%=vname%> class="campo">
	<option value="">- Selecciona de la lista -</option><%
	if not(DSQL2.bof and DSQL2.eof) then
		dSQL2.movefirst
		DO while not dSQL2.eof
		  %><option <%=sele%> value="<%=dSQL2("valor")%>"><%=dSQL2("desc1")%></option><%
	        dSQL2.movenext
	        loop
	end if
%>	
	</select>
<%
else 						'--- tipo 2 = selección visible por radio
	if not(DSQL2.bof and DSQL2.eof) then
		dSQL2.movefirst
		DO while not dSQL2.eof
	%>
	<input type="radio" name="<%=vname%>" class="campo" value="<%=dSQL2("valor")%>">&nbsp;<%=dSQL2("desc1")%>&nbsp;&nbsp;
	<%      dSQL2.movenext
	        loop
	end if
end if
dSQL2.close        
END FUNCTION

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: formulario petición de materiales 2006</title>
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
<SCRIPT LANGUAGE="JavaScript">
<!--
function enviar() 
{
	c1 = document.formulario.direccion.value;
	c2 = document.formulario.localidad.value;
	c3 = document.formulario.provincia.value;
	c4 = document.formulario.cp.value;
	c5 = document.formulario.direccion_materiales.value;
	
	falta = '';
	if (c1=='') { falta = falta+'la dirección'+'\n'; }
	if (c2=='') { falta = falta+'la localidad'+'\n'; }
	if (c3=='') { falta = falta+'la provincia'+'\n'; }
	if (c4=='') { falta = falta+'el código postal'+'\n'; }
	if (c5=='') { falta = falta+'de dónde es la dirección'+'\n'; }
	
	if (falta!='')
		{  alert ('Falta por rellenar:\n\n'+falta) }
	else
		{  document.formulario.submit(); }
}	
// -->
</SCRIPT>	
	
</head>
<body>
<script src="valida_textarea.js"></script>
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
            			<tr class="textmenusup"><td class=textmenusup>Formulario de petición de materiales</td>
          		</table>
			</div>
			
			<% if session("id_ecogente")<>"" then %>
			<div class="textsubmenu" id="submenusup<% response.write (seccion) %>">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
<%            				sql = "SELECT nombre,apellidos,sexo FROM ECOINFORMAS_GENTE WHERE idgente="&session("id_ecogente")
			   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	        set objRecordset = OBJConnection.Execute(sql)
		   	   	        usuario = objRecordset("nombre")&" "&objRecordset("apellidos")
		   	   	        usuario_sexo = "o"
		   	   	        if objRecordset("sexo")=75 then usuario_sexo = "a"
%>
            			<tr><td align="right">Usuari<%=usuario_sexo%> identificad<%=usuario_sexo%>:&nbsp;<%=usuario%></td></tr>
          		</table>
			</div>
       			<% end if %>
			
			<div id="texto">
			
				
				<div class="texto">
<!--- formulario -->				
				
<form method="POST" name="formulario" action="formulario_materiales2006_grabar.asp">

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td class="titulo" align="center" colspan=2>Formulario para solicitar materiales de ECOinformas 2006</td></tr> 
<tr><td class="texto" align="justify" colspan=2>Si perteneces a los <a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=557">colectivos a los que va dirigido el proyecto</a>, por favor rellena este formulario para que podamos atender tu solicitud de envío de materiales.<br><b>Recuerda que el plazo para solicitar el envío de materiales acaba el 26 de enero de 2007.</b></td></tr>
<tr><td class="texto" align="center" colspan=2>&nbsp;</td></tr> 
</table>

<br>&nbsp;

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="2"><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Dirección de envío:</td></tr> </table></td></tr>
<tr>
	<td class=texto align=right valign="middle">Direcci&oacute;n:</td>
	<td class=texto align=left><input type="text" name="direccion" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=texto align=right valign="middle">Localidad:</td>
	<td class=texto align=left><input type="text" name="localidad" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=texto align=right valign="middle">Provincia:</td>
	<td class=texto align=left><% CALL valores("013","provincia","1")%></td>
</tr>
<tr>
	<td class=texto align=right valign="middle">C&oacute;digo postal:</td>
	<td class=texto align=left><input type="text" name="cp" size="5" maxlength="5" class="campo"></td>
</tr>
<tr>
	<td class=texto align=right valign="middle">Indica dónde corresponde esta dirección:</td>
	<td class=texto align=left><% CALL valores("009","direccion_materiales","1")%></td>
</tr>
<tr>
	<td class=texto align=right valign="middle">Puedes hacer alguna anotación que facilite <br>el envío postal (máx. 1000 caracteres):</td>
	<td class=texto align=left><textarea name="comentarios_2006" cols="48" rows="3" class="campo" OnKeyDown="return checkMaxLength(this, event,1000)" OnSelect="storeSelection(this)"></textarea></td>
</tr>
</table>
<br>&nbsp;
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td class="subtitulo" colspan="3" align="left">Marca aquellos materiales que est&aacute;s interesado/a en recibir:</td></tr>
<tr><td>&nbsp;</td><td><strong>Código</strong></td><td><strong>Nombre</strong></td></tr>
<tr><td class="celda"><input type="checkbox" name="cdrom" value="1" class="campo">&nbsp;</td><td class="celda">CD-ROM</td><td class="celda">"acTÚa en tu empresa", CD-ROM interactivo para la mejora ambiental de las PyMES</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP01_2006" value="1" class="campo">&nbsp;</td><td class="celda">EGP01</td><td class="celda">La participación de los trabajadores en la protección del medio ambiente en el centro de trabajo</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE01_2006" value="1" class="campo">&nbsp;</td><td class="celda">AE01</td><td class="celda">Afectación y cumplimiento de la normativa Seveso en la industria española</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE02_2006" value="1" class="campo">&nbsp;</td><td class="celda">AE02</td><td class="celda">Estudio sobre la participación de los trabajadores en el proceso de la Autorización Ambiental Integrada</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE03_2006" value="1" class="campo">&nbsp;</td><td class="celda">AE03</td><td class="celda">Incendios forestales: impacto sobre el medio ambiente, la economía y el empleo</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE04_2006" value="1" class="campo">&nbsp;</td><td class="celda">AE04</td><td class="celda">Estudio del empleo en PyME en el sector de las energías renovables en España</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE05_2006" value="1" class="campo">&nbsp;</td><td class="celda">AE05</td><td class="celda">Estudio sobre la generación y minimización de residuos en el ámbito de las PyMEs</td></tr>
</table>

<br>&nbsp;

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td class=texto align=center>
   <input type="button" value="ENVIAR DATOS" name="comprobar" class="boton" onclick="enviar()">
</td></tr>
<tr><td class=texto align=center>
(*) Los datos que nos facilites serán incorporados a un fichero bajo titularidad de ISTAS. La finalidad del tratamiento de sus datos la constituye la posibilidad de difusión por correo electrónico y ordinario de información y materiales de ECOinformas; la promoción de la salud laboral y la protección del medio ambiente a través de la remisión de información sobre los productos editoriales y actividades de ISTAS; auditoría por parte de la Fundación Biodiversidad que se compromete a su vez a cumplir la Ley Orgánica de Protección de Datos de carácter Personal (LOPD). Para más información: <a href="http://www.istas.net/ecoinformas/index.asp?idpagina=558" target="_blank">política de privacidad.</a>
</td></tr>
   
</table>
</form>

<!-- fin formulario -->				
				</div>
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
