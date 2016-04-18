<!--#include file="EliminaInyeccionSQL.asp"-->
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"




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
<title>ECOinformas: formulario único</title>
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
	c01 = formulario.nombre.value;
	c02 = formulario.apellidos.value;
	c03 = formulario.fec_nac.value;
	c05 = formulario.seg_social.value;       
	c09 = formulario.dni.value;
	c15 = formulario.direccion.value;
	c16 = formulario.localidad.value;
	c18 = formulario.cp.value;
	c19 = formulario.telefono.value; 
	c20 = formulario.movil.value;
	c21 = formulario.fax.value;
	c22 = formulario.email.value;
	c23 = formulario.empresa.value; 
	c24 = formulario.cif.value;
	c25 = formulario.razon_social.value;
	c27 = formulario.emp_direccion.value;
	c28 = formulario.emp_localidad.value;
	c30 = formulario.emp_cp.value;
	c31 = formulario.emp_telefono.value;
	c32 = formulario.emp_movil.value;
	c33 = formulario.emp_fax.value;
	c34 = formulario.emp_email.value; 
	c35 = formulario.emp_web.value;
	c38 = formulario.observaciones.value;
	c39 = formulario.FP02.checked;
	c40 = formulario.FP03.checked;
	c41 = formulario.FP04.checked;
	c42 = formulario.FP01.checked;
	c43 = formulario.FDT01.checked;
	c44 = formulario.SJ01.checked;
	c45 = formulario.SJ02.checked;
	c47 = formulario.FolGen.checked;
	c48 = formulario.FolObs.checked;
	c49 = formulario.SET04.checked;
	c50 = formulario.EGP01.checked;
	c51 = formulario.EGP02.checked;
	c52 = formulario.EGP03.checked;
	c53 = formulario.EGP04.checked;
	c54 = formulario.EGP05.checked;
	c55 = formulario.EGP06.checked;
	c56 = formulario.EGP07.checked;
	c57 = formulario.AE01.checked;
	c58 = formulario.AE02.checked;
	c59 = formulario.AE03.checked;
	c60 = formulario.AE04.checked;
	c61 = formulario.AE05.checked;
	c62 = formulario.AE06.checked;
	
	c10 = formulario.cond_laboral.selectedIndex;
	c11 = formulario.tam_empresa.selectedIndex;
	c12 = formulario.puesto.selectedIndex;
	c13 = formulario.contrato.selectedIndex;
	c14 = formulario.estudios.selectedIndex;
	c17 = formulario.provincia.selectedIndex;
	c26 = formulario.sector.selectedIndex;
	c29 = formulario.emp_provincia.selectedIndex;
	c46 = formulario.direccion_materiales.selectedIndex;

	c04 = false;
	for (i=0;i<formulario.sexo.length;i++)
	{ c04 = ((c04)||(formulario.sexo[i].checked)); }
	c06 = false;
	for (i=0;i<formulario.minusvalia.length;i++)
	{ c06 = ((c06)||(formulario.minusvalia[i].checked)); }
	c07 = false;
	for (i=0;i<formulario.inmigrante.length;i++)
	{ c07 = ((c07)||(formulario.inmigrante[i].checked)); }
	c08 = false;
	for (i=0;i<formulario.cualificacion.length;i++)
	{ c08 = ((c08)||(formulario.cualificacion[i].checked)); }
	c36 = false;
	for (i=0;i<formulario.recibir_info_ecoinformas.length;i++)
	{ c36 = ((c36)||(formulario.recibir_info_ecoinformas[i].checked)); }
	c37 = false;
	for (i=0;i<formulario.recibir_info_istas.length;i++)
	{ c37 = ((c37)||(formulario.recibir_info_istas[i].checked)); }
	
	falta = '';
	if (c01=='') { falta = falta+'nombre'+'\n'; }
	if (c02=='') { falta = falta+'apellidos'+'\n'; }
	if (c03=='') { falta = falta+'fec_nac'+'\n'; }
	if (!c04) { falta = falta+'sexo'+'\n'; }
	if (!c06) { falta = falta+'minusvalia'+'\n'; }
	if (!c07) { falta = falta+'inmigrante'+'\n'; }
	if (!c08) { falta = falta+'cualificacion'+'\n'; }
	if (c10=='0') { falta = falta+'cond_laboral'+'\n'; }
	if (c11=='0') { falta = falta+'tam_empresa'+'\n'; }
	if (c12=='0') { falta = falta+'puesto'+'\n'; }
	if (c13=='0') { falta = falta+'contrato'+'\n'; }
	if (c14=='0') { falta = falta+'estudios'+'\n'; }
	if (c15=='') { falta = falta+'direccion'+'\n'; }
	if (c16=='') { falta = falta+'localidad'+'\n'; }
	if (c17=='0') { falta = falta+'provincia'+'\n'; }
	if (c18=='') { falta = falta+'cp'+'\n'; }
	if (c19=='') { falta = falta+'telefono'+'\n'; }
	//if (c20=='') { falta = falta+'movil'+'\n'; }
	//if (c21=='') { falta = falta+'fax'+'\n'; }
	if (c22=='') { falta = falta+'email'+'\n'; }
	if (c23=='') { falta = falta+'empresa'+'\n'; }
	if (c25=='') { falta = falta+'razon_social'+'\n'; }
	if (c26=='0') { falta = falta+'sector'+'\n'; }
	if (c27=='') { falta = falta+'emp_direccion'+'\n'; }
	if (c28=='') { falta = falta+'emp_localidad'+'\n'; }
	if (c29=='0') { falta = falta+'emp_provincia'+'\n'; }
	if (c30=='') { falta = falta+'emp_cp'+'\n'; }
	if (c31=='') { falta = falta+'emp_telefono'+'\n'; }
	//if (c32=='') { falta = falta+'emp_movil'+'\n'; }
	//if (c33=='') { falta = falta+'emp_fax'+'\n'; }
	//if (c34=='') { falta = falta+'emp_email'+'\n'; }
	//if (c35=='') { falta = falta+'emp_web'+'\n'; }
	if (!c36) { falta = falta+'recibir_info_ecoinformas'+'\n'; }
	if (!c37) { falta = falta+'recibir_info_istas'+'\n'; }
	//if (c38=='') { falta = falta+'observaciones'+'\n'; }
	if ((c41)||(c42)||(c43)||(c44)||(c45))
	{ 
	  if (c05=='') { falta = falta+'seg_social'+'\n'; };
	  if (c09=='') { falta = falta+'dni'+'\n'; };
	  if (c24=='') { falta = falta+'cif'+'\n'; }; 
	}
	
	if ((c47)||(c48)||(c49)||(c50)||(c51)||(c52)||(c53)||(c54)||(c55)||(c56)||(c57)||(c58)||(c59)||(c60)||(c61)||(c62))
	{ 
	  if (c09=='') { falta = falta+'dni'+'\n'; };
	  if (c24=='') { falta = falta+'cif'+'\n'; }; 
	  if (c46=='0') { falta = falta+'direccion_materiales'+'\n'; };
	}
	
	
	if (falta!='')
	{  alert ('Falta por rellenar:\n\n'+falta) }
	else
	{  document.formulario.submit(); }
}

function mostrar(cual)
{
	eval("oculto"+cual+".style.visibility = 'hidden'");
	eval("oculto"+cual+".style.display = 'none'");
	eval("visible"+cual+".style.visibility = 'visible'");
	eval("visible"+cual+".style.display = 'block'");
}

function ocultar(cual)
{
	eval("oculto"+cual+".style.visibility = 'visible'");
	eval("oculto"+cual+".style.display = 'block'");
	eval("visible"+cual+".style.visibility = 'hidden'");
	eval("visible"+cual+".style.display = 'none'");
}

// -->
</SCRIPT>
</head>
<body>
<script src="valida_fecha.js"></script>
<script src="valida_mail.js"></script>
<script src="valida_textarea.js"></script>
<script src="valida_dni.js"></script>

<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
			<div id="encabezado_nuevo1">
			<table width="100%" cellpadding=0 border=0>
			<tr><td width="215" height="78" onClick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="142" height="78" onClick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="166" height="78" onClick="location.href='index.asp?idpagina=549'" style="cursor:hand">&nbsp;</td>
			    <td width="160" height="78" onClick="location.href='index.asp?idpagina=550'" style="cursor:hand">&nbsp;</td>
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
<%            				sql = "SELECT nombre,apellidos,sexo FROM ECOINFORMAS_GENTE WHERE idgente="&session("id_ecogente")
			   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	        set objRecordset = OBJConnection.Execute(sql)
		   	   	        usuario = objRecordset("nombre")&" "&objRecordset("apellidos")
		   	   	        usuario_sexo = "o"
		   	   	        if objRecordset("sexo")=75 then usuario_sexo = "a"
%>
            			<tr><td align="right">No es necesario que rellenes el formulario con tus datos. Usuari<%=usuario_sexo%> identificad<%=usuario_sexo%>:&nbsp;<%=usuario%></td></tr>
          		</table>
			</div>
       			<% end if %>
			
			<div id="texto">
			
				
				<div class="texto">
<!--- formulario -->				
				
<form method="POST" name="formulario" action="formulario2_grabar.asp">

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td class="titulo" align="center" colspan=2>Formulario para participar en ECOinformas</td></tr> 
<tr><td class="texto" align="justify" colspan=2>Por favor rellena este formulario para que podamos atender tu solicitud de acceso libre a toda nuestra página web y/o participar en los cursos de formación y/o solicitar materiales (*). Si perteneces a los colectivos a los que va dirigido este proyecto, recibirás por correo electrónico una clave y una contraseña que te permiten entrar en la página web y aprovechar los servicios y materiales que te ofrecemos, incluido el servicio de asesoramiento directo del observatorio medioambiental. Además recibirás un boletín electrónico con novedades en la página web y noticias de ECOinformas.</td></tr>
<tr><td class="texto" align="center" colspan=2>&nbsp;</td></tr> 
</table>

<br>&nbsp;

<div id="oculto1" style="overflow: auto; visibility: hidden; display: none"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Datos personales</td><td class=texto" align="right"><a onClick="mostrar('1')" style="cursor:hand">Mostrar campos [+]</a></td></tr></table></td></tr>
</table>
</div>

<div id="visible1" style="overflow: auto; visibility: visible; display: block"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="2"><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Datos personales</td><td class=texto" align="right"><a onClick="ocultar('1')" style="cursor:hand">Ocultar campos [-]</a></td></tr> </table></td></tr>
<tr>
	<td class=celdacentro align=right valign="middle">Nombre:</td>
	<td class=celdacentro align=left><input type="text" name="nombre" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Apellidos:</td>
	<td class=celdacentro align=left><input type="text" name="apellidos" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Fecha nacimiento:</td>
	<td class=celdacentro align=left><input type="text" name="fec_nac" size="11" maxlength="10" class="campo" OnBlur='valida_fecha(this)'></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Sexo:</td>
	<td class=celdacentro align=left><% CALL valores("001","sexo","2")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Núm. Seguridad Social:<br>&nbsp;</td>
	<td class=celdacentro align=left><input type="text" name="seg_social" size="50" maxlength="50" class="campo"><br>Rellénalo sólo en el caso de solicitar la inscripción en cursos/jornadas</td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Minusvalía reconocida:</td>
	<td class=celdacentro align=left><% CALL valores("002","minusvalia","2")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Eres inmigrante:</td>
	<td class=celdacentro align=left><% CALL valores("002","inmigrante","2")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Eres trabajador<br>de baja cualificación:</td>
	<td class=celdacentro align=left><% CALL valores("002","cualificacion","2")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">DNI/NIE:<br>&nbsp;<br>&nbsp;</td>
	<td class=celdacentro align=left><input type="text" name="dni" size="10" maxlength="9" class="campo" OnBlur='valida_dni(this)'>&nbsp;(escribe sólo los números, la letra sale automáticamente)<br>Rellénalo sólo en el caso de solicitar la inscripción en cursos/jornadas <br>o petición de materiales</td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Condición laboral:</td>
	<td class=celdacentro align=left><% CALL valores("003","cond_laboral","1")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Tamaño de tu empresa:</td>
	<td class=celdacentro align=left><% CALL valores("004","tam_empresa","1")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Puesto que desempeñas:</td>
	<td class=celdacentro align=left><% CALL valores("005","puesto","1")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Tipo de contrataci&oacute;n:</td>
	<td class=celdacentro align=left><% CALL valores("006","contrato","1")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Datos acad&eacute;micos:</td>
	<td class=celdacentro align=left><% CALL valores("007","estudios","1")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Direcci&oacute;n (residencia habitual):</td>
	<td class=celdacentro align=left><input type="text" name="direccion" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Localidad:</td>
	<td class=celdacentro align=left><input type="text" name="localidad" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Provincia:</td>
	<td class=celdacentro align=left><% CALL valores("013","provincia","1")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">C&oacute;digo postal:</td>
	<td class=celdacentro align=left><input type="text" name="cp" size="5" maxlength="5" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Tel&eacute;fono:</td>
	<td class=celdacentro align=left><input type="text" name="telefono" size="50" maxlength="50" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">M&oacute;vil:</td>
	<td class=celdacentro align=left><input type="text" name="movil" size="50" maxlength="50" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Fax:</td>
	<td class=celdacentro align=left><input type="text" name="fax" size="50" maxlength="50" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Email:</td>
	<td class=celdacentro align=left><input type="text" name="email" size="50" maxlength="200" class="campo" ONBLUR='valida_mail(this)'></td>
</tr>
</table>
</div>

<br>&nbsp;

<div id="oculto2" style="overflow: auto; visibility: hidden; display: none"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Datos de la empresa en que trabajas</td><td class=texto" align="right"><a onClick="mostrar('2')" style="cursor:hand">Mostrar campos [+]</a></td></tr></table></td></tr>
</table>
</div>

<div id="visible2" style="overflow: auto; visibility: visible; display: block"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="2"><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Datos de la empresa en que trabajas</td><td class=texto" align="right"><a onClick="ocultar('2')" style="cursor:hand">Ocultar campos [-]</a></td></tr> </table></td></tr>
<tr>
	<td class=celdacentro align=right valign="middle">Nombre empresa:</td>
	<td class=celdacentro align=left><input type="text" name="empresa" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">CIF:<br>&nbsp;<br>&nbsp;</td>
	<td class=celdacentro align=left><input type="text" name="cif" size="10" maxlength="10" class="campo"><br>Rellénalo sólo en el caso de solicitar la inscripción en cursos/jornadas <br>o petición de materiales</td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Raz&oacute;n social:</td>
	<td class=celdacentro align=left><input type="text" name="razon_social" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Sector:</td>
	<td class=celdacentro align=left><% CALL valores("008","sector","1")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Direcci&oacute;n:</td>
	<td class=celdacentro align=left><input type="text" name="emp_direccion" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Localidad:</td>
	<td class=celdacentro align=left><input type="text" name="emp_localidad" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Provincia:</td>
	<td class=celdacentro align=left><% CALL valores("013","emp_provincia","1")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Código postal:</td>
	<td class=celdacentro align=left><input type="text" name="emp_cp" size="5" maxlength="5" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Teléfono:</td>
	<td class=celdacentro align=left><input type="text" name="emp_telefono" size="50" maxlength="50" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">M&oacute;vil:</td>
	<td class=celdacentro align=left><input type="text" name="emp_movil" size="50" maxlength="50" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Fax:</td>
	<td class=celdacentro align=left><input type="text" name="emp_fax" size="50" maxlength="50" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Email:</td>
	<td class=celdacentro align=left><input type="text" name="emp_email" size="50" maxlength="200" class="campo" ONBLUR='valida_mail(this)'></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">P&aacute;gina web:</td>
	<td class=celdacentro align=left><input type="text" name="emp_web" size="50" maxlength="50" class="campo" value="http://"></td>
</tr>
</table>
</div>

<br>&nbsp;

<div id="oculto3" style="overflow: auto; visibility: hidden; display: none"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Otros datos/opciones</td><td class=texto" align="right"><a onClick="mostrar('3')" style="cursor:hand">Mostrar campos [+]</a></td></tr></table></td></tr>
</table>
</div>

<div id="visible3" style="overflow: auto; visibility: visible; display: block"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="2"><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Otros datos/opciones</td><td class=texto" align="right"><a onClick="ocultar('3')" style="cursor:hand">Ocultar campos [-]</a></td></tr> </table></td></tr>
<tr><td class="celda" align="left" colspan=2>Si te interesa además recibir información periódica sobre medio ambiente y salud laboral, por favor <br>completa los siguientes campos también:</td></tr>
<tr><td class="celda" align="left" colspan=2>Según la normativa vigente en protección de datos personales, necesitamos su consentimiento explícito para poder utilizar los datos.</td></tr>
<tr>
	<td class=celdacentro align=right valign="middle">Quiero recibir informaci&oacute;n peri&oacute;dica sobre<br>las actividades y publicaciones de ECOinformas:</td>
	<td class=celdacentro align=left><% CALL valores("002","recibir_info_ecoinformas","2")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Me interesa recibir informaci&oacute;n peri&oacute;dica sobre<br>temas de medio ambiente y salud laboral, distribuida por ISTAS:</td>
	<td class=celdacentro align=left><% CALL valores("002","recibir_info_istas","2")%></td>
</tr>
</table>
</div>

<br>&nbsp;

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="2"><table width="90%" align="center"><tr><td class="subtitulo" align="left">Si quieres anotar cualquier observación junto al formulario</td></tr></table></td></tr>
<tr>
	<td class=celdacentro align=right valign="top">Observaciones:</td>
	<td class=celdacentro align=left><textarea name="observaciones" cols=50 rows=5 class="campo" OnKeyDown="return checkMaxLength(this, event,3000)" OnSelect="storeSelection(this)"></textarea></td>
</tr>
</table>

<br>&nbsp;

<div id="oculto4" style="overflow: auto; visibility: hidden; display: none"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Solicitud inscripci&oacute;n en cursos y jornadas</td><td class=texto" align="right"><a onClick="mostrar('4')" style="cursor:hand">Mostrar campos [+]</a></td></tr></table></td></tr>
</table>
</div>

<div id="visible4" style="overflow: auto; visibility: visible; display: block"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="10"><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Solicitud inscripci&oacute;n en cursos y jornadas</td><td class=texto" align="right"><a onClick="ocultar('4')" style="cursor:hand">Ocultar campos [-]</a></td></tr> </table></td></tr>
<tr><td class="celda" align="left" colspan=10>Marca aquellos cursos/jornadas donde quieres inscribirte:</td></tr>
<tr><td>&nbsp;</td><td><strong>Código</strong></td><td><strong>Nombre</strong></td><td><strong>Lugar de realización</strong></td><td><strong>Fecha</strong></td><td><strong>Horario</strong></td><td><strong>Inscripciones</strong></td></tr>
<tr><td class="celda"><input type="checkbox" name="FP02" value="1" class="campo">&nbsp;</td><td class="celda">FP02</td><td class="celda"><a href=index.asp?idpagina=581 target="_blank">Medio Ambiente en la empresa </a></td><td class="celda">MADRID: U.S.M.R. de CC.OO. FITEQA C/ Lope de Vega 38, 2ª planta </td><td class="celda">06/10/2005</td><td class="celda">9:00 a 14:00</td><td class="celda">15 al 27 de septiembre</td></tr>
<tr><td class="celda"><input type="checkbox" name="FP03" value="1" class="campo">&nbsp;</td><td class="celda">FP03</td><td class="celda"><a href=index.asp?idpagina=581 target="_blank">Medio Ambiente en la empresa</a></td><td class="celda">MANLLEU (BARCELONA): U.C. Osona de CC.OO. Catalunya Pz. del Mercat, 100 </td><td class="celda">18/10/2005</td><td class="celda">9:00 a 14:00 </td><td class="celda">15 de septiembre al 4 de octubre</td></tr>
<tr><td class="celda"><input type="checkbox" name="FP04" value="1" class="campo">&nbsp;</td><td class="celda">FP04</td><td class="celda"><a href=index.asp?idpagina=581 target="_blank">Medio Ambiente en la empresa</a></td><td class="celda">MURCIA: CC.OO. Región de Murcia C/ Corbalán, 4 </td><td class="celda">16/11/2005</td><td class="celda">9:00 a 14:00 </td><td class="celda">15 de septiembre al 2 de noviembre</td></tr>
<tr><td class="celda"><input type="checkbox" name="FP01" value="1" class="campo">&nbsp;</td><td class="celda">FP01</td><td class="celda"><a href=index.asp?idpagina=581 target="_blank">Medio Ambiente en la empresa </a></td><td class="celda">MADRID: U.S.M.R. de CC.OO. Fed. Minero-metalúrgica. C/ Lope de Vega 38 6ª Planta </td><td class="celda">15/12/2005</td><td class="celda">9:00 a 14:00</td><td class="celda">15 de septiembre al 30 de noviembre</td></tr>
<tr><td class="celda"><input type="checkbox" name="FDT01" value="1" class="campo">&nbsp;</td><td class="celda">FDT01<br>FDT02<br>FDT03</td><td class="celda"><a href=index.asp?idpagina=582 target="_blank">Curso on-line sobre “Medio Ambiente, Salud y Desarrollo Sostenible”</a></td><td class="celda">Plataforma de formación a distancia en Internet</td><td class="celda">a partir de 20/10/2005</td><td class="celda">20 horas en total</td><td class="celda">del 15 de septiembre al 7 de octubre y después en función de disponiblidad de plazas</td></tr>
<tr><td class="celda"><input type="checkbox" name="SJ01" value="1" class="campo">&nbsp;</td><td class="celda">SJ01</td><td class="celda"><a href=index.asp?idpagina=590 target="_blank">Jornada de Prevención de la Contaminación</a></td><td class="celda">ZARAGOZA: Centro Cultural Joaquín Roncal (Obras Social de la Caja de Ahorros de la Inmaculada)</td><td class="celda">03/11/2005</td><td class="celda">9:30 a 14:30</td><td class="celda">15 de septiembre al 28 de octubre</td></tr>
<tr><td class="celda"><input type="checkbox" name="SJ02" value="1" class="campo">&nbsp;</td><td class="celda">SJ02</td><td class="celda"><a href=index.asp?idpagina=590 target="_blank">Jornada de Prevención de la Contaminación </a></td><td class="celda">VALENCIA: lugar exacto se comunicará en breve</td><td class="celda">fecha por confirmar</td><td class="celda">9:00 a 14:00</td><td class="celda">por confirmar</td></tr>
</table>
</div>

<br>&nbsp;

<div id="oculto5" style="overflow: auto; visibility: hidden; display: none"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="50%">Solicitud para recibir materiales</td><td class=texto" align="right"><a onClick="mostrar('5')" style="cursor:hand">Mostrar campos [+]</a></td></tr></table></td></tr>
</table>
</div>

<div id="visible5" style="overflow: auto; visibility: visible; display: block"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="10"><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="50%">Solicitud para recibir materiales</td><td class=texto" align="right"><a onClick="ocultar('5')" style="cursor:hand">Ocultar campos [-]</a></td></tr> </table></td></tr>
<tr><td class="celda" colspan="3" align="left">
	<table>
	<tr>
	<td class=texto align=right valign="middle">Domicilio para enviar los materiales:</td>
	<td class=texto align=left><% CALL valores("009","direccion_materiales","1")%></td>
	</tr>
	</table>
    </td>
</tr>
<tr><td class="celda" colspan="3" align="left">Marca aquellos materiales que est&aacute;s interesado/a en recibir:</td></tr>
<tr><td>&nbsp;</td><td><strong>Código</strong></td><td><strong>Nombre</strong></td></tr>
<tr><td class="celda"><input type="checkbox" name="FolGen" value="1" class="campo">&nbsp;</td><td class="celda">FolGen</td><td class="celda">Folleto general sobre ECOinformas</td></tr>
<tr><td class="celda"><input type="checkbox" name="FolObs" value="1" class="campo">&nbsp;</td><td class="celda">FolObs</td><td class="celda">Folleto general sobre el Observatorio Medioambiental</td></tr>
<tr><td class="celda"><input type="checkbox" name="SET04" value="1" class="campo">&nbsp;</td><td class="celda">SET04</td><td class="celda">Vídeo informativo: El riesgo químico en mi empresa</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP01" value="1" class="campo">&nbsp;</td><td class="celda">EGP01</td><td class="celda">Guía para la sustitución de sustancias peligrosas en las empresas</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP02" value="1" class="campo">&nbsp;</td><td class="celda">EGP02</td><td class="celda">Guía de control y gestión de residuos peligrosos</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP03" value="1" class="campo">&nbsp;</td><td class="celda">EGP03</td><td class="celda">Guía de gestión y control de emisiones</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP04" value="1" class="campo">&nbsp;</td><td class="celda">EGP04</td><td class="celda">Guía de ahorro y eficiencia energética</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP05" value="1" class="campo">&nbsp;</td><td class="celda">EGP05</td><td class="celda">Guía de buenas prácticas para la minimización de residuos, emisiones y vertidos</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP06" value="1" class="campo">&nbsp;</td><td class="celda">EGP06</td><td class="celda">Guía de ahorro de agua</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP07" value="1" class="campo">&nbsp;</td><td class="celda">EGP07</td><td class="celda">Guía de gestión y control de vertidos</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE01" value="1" class="campo">&nbsp;</td><td class="celda">AE01</td><td class="celda">Incidencia de la Aplicación del Protocolo de Kioto y del Plan Nacional de Asignación en las PYMEs españolas del sector industrial</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE02" value="1" class="campo">&nbsp;</td><td class="celda">AE02</td><td class="celda">Necesidades de información sobre medio ambiente por parte de trabajadores de PYMEs</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE03" value="1" class="campo">&nbsp;</td><td class="celda">AE03</td><td class="celda">Estudio de las condiciones sociolaborales para la mejora ambiental en el sector cerámico de Bailén</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE04" value="1" class="campo">&nbsp;</td><td class="celda">AE04</td><td class="celda">Prevención del riesgo químico en PYMEs. Fuentes de información y herramientas</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE05" value="1" class="campo">&nbsp;</td><td class="celda">AE05</td><td class="celda">Evaluación del impacto de REACH sobre la salud laboral en PYMEs españolas</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE06" value="1" class="campo">&nbsp;</td><td class="celda">AE06</td><td class="celda">Estudio sobre requisitos legales (estatales y autonómicos) y aspectos ambientales aplicables a PYMEs afectadas por la Ley 16/2002 de Prevención y Control Integrado de la Contaminación (LPCIC)</td></tr>
</table>
</div>

<br>&nbsp;

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td class=texto align=center>
(*) Los datos que nos facilites serán incorporados a un fichero bajo titularidad de ISTAS. La finalidad del tratamiento de sus datos la constituye la posibilidad de difusión por correo electrónico y ordinario de información y materiales de ECOinformas; la promoción de la salud laboral y la protección del medio ambiente a través de la remisión de información sobre los productos editoriales y actividades de ISTAS; auditoría por parte de la Fundación Biodiversidad que se compromete a su vez a cumplir la Ley Orgánica de Protección de Datos de carácter Personal (LOPD). Para más información: <a href="http://www.istas.net/ecoinformas/index.asp?idpagina=558" target="_blank">política de privacidad.</a>
</td></tr>
<tr><td class=texto align=center>
   <input type="button" value="ENVIAR DATOS" name="comprobar" class="boton" onClick="enviar()">
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
