<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

	'usuario = request.cookies("webistas")

if session("id_ecogente")<>"" then response.redirect "formulario_identificado.asp"


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
else 						'--- tipo 2 = selecci�n visible por radio
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
<title>ECOinformas: formulario �nico</title>
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
<SCRIPT LANGUAGE="JavaScript">
<!--
function enviar() 
{
	c01 = document.formulario.nombre.value;
	c02 = document.formulario.apellidos.value;
	c03 = document.formulario.fec_nac.value;
	//c05 = document.formulario.seg_social.value;       
	//c09 = document.formulario.dni.value;
	//c15 = document.formulario.direccion.value;
	//c16 = document.formulario.localidad.value;
	//c18 = document.formulario.cp.value;
	c19 = document.formulario.telefono.value; 
	//c20 = document.formulario.movil.value;
	//c21 = document.formulario.fax.value;
	c22 = document.formulario.email.value;
	c23 = document.formulario.empresa.value; 
	//c24 = document.formulario.cif.value;
	//c25 = document.formulario.razon_social.value;
	//c27 = document.formulario.emp_direccion.value;
	//c28 = document.formulario.emp_localidad.value;
	//c30 = document.formulario.emp_cp.value;
	//c31 = document.formulario.emp_telefono.value;
	//c32 = document.formulario.emp_movil.value;
	//c33 = document.formulario.emp_fax.value;
	//c34 = document.formulario.emp_email.value; 
	//c35 = document.formulario.emp_web.value;
	c38 = document.formulario.observaciones.value;
	//c39 = document.formulario.FP02.checked;
	//c40 = document.formulario.FP03.checked;
	//c41 = document.formulario.FP04.checked;
	//c42 = document.formulario.FP01.checked;
	//c43 = document.formulario.FDT01.checked;
	//c44 = document.formulario.SJ01.checked;
	//c45 = document.formulario.SJ02.checked;
	//c47 = document.formulario.FolGen.checked;
	//c48 = document.formulario.FolObs.checked;
	//c49 = document.formulario.SET04.checked;
	//c50 = document.formulario.EGP01.checked;
	//c51 = document.formulario.EGP02.checked;
	//c52 = document.formulario.EGP03.checked;
	//c53 = document.formulario.EGP04.checked;
	//c54 = document.formulario.EGP05.checked;
	//c55 = document.formulario.EGP06.checked;
	//c56 = document.formulario.EGP07.checked;
	//c57 = document.formulario.AE01.checked;
	//c58 = document.formulario.AE02.checked;
	//c59 = document.formulario.AE03.checked;
	//c60 = document.formulario.AE04.checked;
	//c61 = document.formulario.AE05.checked;
	//c62 = document.formulario.AE06.checked;
	//c63 = document.formulario.SEP01.checked;
	//c64 = document.formulario.SEP02.checked;
	//c65 = document.formulario.SEP03.checked;
	
	c10 = document.formulario.cond_laboral.selectedIndex;
	c11 = document.formulario.tam_empresa.selectedIndex;
	//c12 = document.formulario.puesto.selectedIndex;
	//c13 = document.formulario.contrato.selectedIndex;
	//c14 = document.formulario.estudios.selectedIndex;
	//c17 = document.formulario.provincia.selectedIndex;
	//c26 = document.formulario.sector.selectedIndex;
	c29 = document.formulario.emp_provincia.selectedIndex;
	//c46 = document.formulario.direccion_materiales.selectedIndex;

	c04 = false;
	for (i=0;i<document.formulario.sexo.length;i++)
	{ c04 = ((c04)||(document.formulario.sexo[i].checked)); }
	c06 = false;
	for (i=0;i<document.formulario.minusvalia.length;i++)
	{ c06 = ((c06)||(document.formulario.minusvalia[i].checked)); }
	c07 = false;
	for (i=0;i<document.formulario.inmigrante.length;i++)
	{ c07 = ((c07)||(document.formulario.inmigrante[i].checked)); }
	c08 = false;
	for (i=0;i<document.formulario.cualificacion.length;i++)
	{ c08 = ((c08)||(document.formulario.cualificacion[i].checked)); }
	c36 = false;
	for (i=0;i<document.formulario.recibir_info_ecoinformas.length;i++)
	{ c36 = ((c36)||(document.formulario.recibir_info_ecoinformas[i].checked)); }
	c37 = false;
	for (i=0;i<document.formulario.recibir_info_istas.length;i++)
	{ c37 = ((c37)||(document.formulario.recibir_info_istas[i].checked)); }
	
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
	//if (c12=='0') { falta = falta+'puesto'+'\n'; }
	//if (c13=='0') { falta = falta+'contrato'+'\n'; }
	//if (c14=='0') { falta = falta+'estudios'+'\n'; }
	if (c19=='') { falta = falta+'telefono'+'\n'; }
	//if (c20=='') { falta = falta+'movil'+'\n'; }
	//if (c21=='') { falta = falta+'fax'+'\n'; }
	if (c22=='') { falta = falta+'email'+'\n'; }
	if (c23=='') { falta = falta+'empresa'+'\n'; }
	//if (c25=='') { falta = falta+'razon_social'+'\n'; }
	//if (c26=='0') { falta = falta+'sector'+'\n'; }
	//if (c27=='') { falta = falta+'emp_direccion'+'\n'; }
	//if (c28=='') { falta = falta+'emp_localidad'+'\n'; }
	if (c29=='0') { falta = falta+'emp_provincia'+'\n'; }
	//if (c30=='') { falta = falta+'emp_cp'+'\n'; }
	//if (c31=='') { falta = falta+'emp_telefono'+'\n'; }
	//if (c32=='') { falta = falta+'emp_movil'+'\n'; }
	//if (c33=='') { falta = falta+'emp_fax'+'\n'; }
	//if (c34=='') { falta = falta+'emp_email'+'\n'; }
	//if (c35=='') { falta = falta+'emp_web'+'\n'; }
	//if (!c36) { falta = falta+'recibir_info_ecoinformas'+'\n'; }
	//if (!c37) { falta = falta+'recibir_info_istas'+'\n'; }
	//if (c38=='') { falta = falta+'observaciones'+'\n'; }
	//if ((c41)||(c42)||(c43)||(c44)||(c45))
	//{ 
	//  if (c05=='') { falta = falta+'seg_social'+'\n'; };
	//  if (c09=='') { falta = falta+'dni'+'\n'; };
	//  if (c24=='') { falta = falta+'cif'+'\n'; }; 
	//}
	
	//if ((c49)||(c50)||(c51)||(c52)||(c53)||(c54)||(c55)||(c56)||(c57)||(c58)||(c59)||(c60)||(c61)||(c62)||(c63)||(c64)||(c65))
	//{ 
	  //if (c09=='') { falta = falta+'dni'+'\n'; };
	  //if (c24=='') { falta = falta+'cif'+'\n'; }; 
	  //if (c46=='0') { falta = falta+'direccion_materiales'+'\n'; };
  	  //if (c15=='') { falta = falta+'direccion'+'\n'; }
	  //if (c16=='') { falta = falta+'localidad'+'\n'; }
	  //if (c17=='0') { falta = falta+'provincia'+'\n'; }
	  //if (c18=='') { falta = falta+'cp'+'\n'; }

	//}
	
	
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
            			<tr class="textmenusup"><td class=textmenusup>Formulario de alta</td>
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
				
<form method="POST" name="formulario" action="formulario4_grabar.asp">

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td class="titulo" align="center" colspan=2>Formulario para participar en ECOinformas</td></tr> 
<tr><td class="texto" align="justify" colspan=2>Por favor rellena este formulario para que podamos atender tu solicitud de acceso libre a toda nuestra p�gina web*. Autom�ticamente obtendr�s una clave y una contrase�a que te permiten entrar en la web ECOinformas y aprovechar los materiales y herramientas que te ofrecemos. Adem�s recibir�s un bolet�n electr�nico con novedades en la p�gina web y noticias de ECOinformas. Si perteneces a los <a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=557" target="_blank">colectivos a los que va dirigido este proyecto</a>, <b>podr�s solicitar que te enviemos de forma gratuita los materiales impresos que te interesen o bien inscribirte en los cursos gratuitos a distancia.</b></td></tr>
<tr><td class="texto" align="justify" colspan=2>Si ya eres usuario(a) de ECOinformas y quieres inscribirte para un curso, aseg�rate de que est�s identificado(a) con tu clave y contrase�a antes de hacerlo. De esta manera, s�lo tendr�s que rellenar algunos campos adicionales del formulario y evitar�s registrar todos tus datos de nuevo.</td></tr>
<tr><td class="texto" align="center" colspan=2>&nbsp;</td></tr> 
</table>

<br>&nbsp;

<div id="oculto1" style="overflow: auto; visibility: hidden; display: none"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Datos personales</td><td class=texto" align="right"><a onclick="mostrar('1')" style="cursor:hand">Mostrar campos [+]</a></td></tr></table></td></tr>
</table>
</div>

<div id="visible1" style="overflow: auto; visibility: visible; display: block"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="2"><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Datos personales</td><td class=texto" align="right"><a onclick="ocultar('1')" style="cursor:hand">Ocultar campos [-]</a></td></tr> </table></td></tr>
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
	<td class=celdacentro align=right valign="middle">Minusval�a reconocida:</td>
	<td class=celdacentro align=left><% CALL valores("002","minusvalia","2")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Inmigrante:</td>
	<td class=celdacentro align=left><% CALL valores("002","inmigrante","2")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Trabajador/a de baja cualificaci�n:</td>
	<td class=celdacentro align=left><% CALL valores("002","cualificacion","2")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Tel&eacute;fono de contacto:</td>
	<td class=celdacentro align=left><input type="text" name="telefono" size="50" maxlength="50" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Email de contacto:</td>
	<td class=celdacentro align=left><input type="text" name="email" size="50" maxlength="200" class="campo" ONBLUR='valida_mail(this)'></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Condici�n laboral:</td>
	<td class=celdacentro align=left><% CALL valores("003","cond_laboral","1")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Nombre empresa:</td>
	<td class=celdacentro align=left><input type="text" name="empresa" size="50" maxlength="200" class="campo"></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Tama�o empresa:</td>
	<td class=celdacentro align=left><% CALL valores("004","tam_empresa","1")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Provincia empresa:</td>
	<td class=celdacentro align=left><% CALL valores("013","emp_provincia","1")%></td>
</tr>
</table>
</div>

<br>&nbsp;

<% if 1=0 then %>
<div id="oculto2" style="overflow: auto; visibility: visible; display: block"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td><table width="90%" align="center"><tr><td class="texto" align="left" width="75%">Si perteneces a alguno de los <a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=557" target="_blank">colectivos a los que va dirigido este proyecto</a> y quieres recibir materiales impresos por correo postal de forma gratuita rellena los siguientes campos adicionales</td><td class=texto" align="right"><a onclick="mostrar('2')" style="cursor:hand">Mostrar campos [+]</a></td></tr></table></td></tr>
</table>
</div>

<div id="visible2" style="overflow: auto; visibility: hidden; display: none"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="10"><table width="90%" align="center"><tr><td class="texto" align="left" width="75%">Si perteneces a alguno de los <a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=557" target="_blank">colectivos a los que va dirigido este proyecto</a> y quieres recibir materiales impresos por correo postal de forma gratuita rellena los siguientes campos adicionales</td><td class=texto" align="right"><a onclick="ocultar('2')" style="cursor:hand">Ocultar campos [-]</a></td></tr> </table></td></tr>
<tr><td class="celda" colspan="3" align="left">
	<table>
	<tr>
	<td class=texto align=right valign="middle">Domicilio para enviar los materiales:</td>
	<td class=texto align=left><% CALL valores("009","direccion_materiales","1")%></td>
	</tr>
	</table>
    </td>
</tr>
<tr><td class="celda" colspan="3" align="left">
<table width="100%">
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
</table>
</td></tr>
<tr><td class="celda" colspan="3" align="left">Marca aquellos materiales que est&aacute;s interesado/a en recibir:</td></tr>
<tr><td>&nbsp;</td><td><strong>C�digo</strong></td><td><strong>Nombre</strong></td></tr>
<tr><td class="celda"><input type="checkbox" name="SEP01" value="1" class="campo">&nbsp;</td><td class="celda">SEP01</td><td class="celda">Cartel de sensibilizaci�n en riesgo qu�mico</td></tr>
<tr><td class="celda"><input type="checkbox" name="SEP02" value="1" class="campo">&nbsp;</td><td class="celda">SEP02</td><td class="celda">Folleto de sensibilizaci�n en riesgo qu�mico</td></tr>
<tr><td class="celda"><input type="checkbox" name="SEP03" value="1" class="campo">&nbsp;</td><td class="celda">SEP03</td><td class="celda">Ficha de identificaci�n del riesgo qu�mico en el lugar de trabajo</td></tr>
<tr><td class="celda"><input type="checkbox" name="SET04" value="1" class="campo">&nbsp;</td><td class="celda">SET04</td><td class="celda">V�deo informativo: "Riesgo qu�mico: �conoces lo que usas?"</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP01" value="1" class="campo">&nbsp;</td><td class="celda">EGP01</td><td class="celda">Gu�a para la sustituci�n de sustancias peligrosas en las empresas</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP02" value="1" class="campo">&nbsp;</td><td class="celda">EGP02</td><td class="celda">Gu�a de control y gesti�n de residuos peligrosos</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP03" value="1" class="campo">&nbsp;</td><td class="celda">EGP03</td><td class="celda">Gu�a de gesti�n y control de emisiones</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP04" value="1" class="campo">&nbsp;</td><td class="celda">EGP04</td><td class="celda">Gu�a de ahorro y eficiencia energ�tica</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP05" value="1" class="campo">&nbsp;</td><td class="celda">EGP05</td><td class="celda">Las buenas pr�cticas para la mejora ambiental en la empresa</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP06" value="1" class="campo">&nbsp;</td><td class="celda">EGP06</td><td class="celda">Gu�a de ahorro de agua</td></tr>
<tr><td class="celda"><input type="checkbox" name="EGP07" value="1" class="campo">&nbsp;</td><td class="celda">EGP07</td><td class="celda">Gu�a de gesti�n y control de vertidos</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE01" value="1" class="campo">&nbsp;</td><td class="celda">AE01</td><td class="celda">Incidencia de la Aplicaci�n del Protocolo de Kioto y del Plan Nacional de Asignaci�n en las PYMEs espa�olas del sector industrial</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE02" value="1" class="campo">&nbsp;</td><td class="celda">AE02</td><td class="celda">Necesidades de informaci�n sobre medio ambiente por parte de trabajadores de PYMEs</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE03" value="1" class="campo">&nbsp;</td><td class="celda">AE03</td><td class="celda">Estudio de las condiciones sociolaborales para la mejora ambiental en el sector cer�mico de Bail�n</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE04" value="1" class="campo">&nbsp;</td><td class="celda">AE04</td><td class="celda">Prevenci�n del riesgo qu�mico en PYME: fuentes de informaci�n y herramientas</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE05" value="1" class="campo">&nbsp;</td><td class="celda">AE05</td><td class="celda">Evaluaci�n del impacto de REACH sobre la salud laboral en PYMEs espa�olas</td></tr>
<tr><td class="celda"><input type="checkbox" name="AE06" value="1" class="campo">&nbsp;</td><td class="celda">AE06</td><td class="celda">Estudio sobre requisitos legales (estatales y auton�micos) y aspectos ambientales aplicables a PYMEs afectadas por la Ley 16/2002 de Prevenci�n y Control Integrado de la Contaminaci�n (LPCIC)</td></tr>
</table>
</div>

<br>&nbsp;
<% end if %>

<% if 1=0 then %>
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="10"><table width="90%" align="center"><tr><td class="texto" align="left" width="75%">Si quieres solicitar la participaci�n en uno de los siguientes cursos, m�rcalo:</td></tr> </table></td></tr>
<tr><td>&nbsp;</td><td><strong>C�digo</strong></td><td><strong>Curso on-line</strong></td><td><strong>Fecha</strong></td><td><strong>Inscripciones</strong></td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT01" value="1" class="campo">&nbsp;</td><td class="celda">FDT01</td><td class="celda">Medio Ambiente, Salud y Desarrollo Sostenible</td><td class="celda">22/05/2006</td><td class="celda">hasta 18 de mayo</td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT02" value="1" class="campo">&nbsp;</td><td class="celda">FDT02</td><td class="celda">Medio Ambiente, Salud y Desarrollo Sostenible</td><td class="celda">29/05/2006</td><td class="celda">hasta 25 de mayo</td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT03" value="1" class="campo">&nbsp;</td><td class="celda">FDT03</td><td class="celda">Medio ambiente y actividades productivas: Pr�cticas sostenibles en agua, energ�a, residuos y emisiones</td><td class="celda">18/09/2006</td><td class="celda">hasta 14 de septiembre</td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT04" value="1" class="campo">&nbsp;</td><td class="celda">FDT04</td><td class="celda">Medio ambiente y actividades productivas: Pr�cticas sostenibles en agua, energ�a, residuos y emisiones</td><td class="celda">25/09/2006</td><td class="celda">hasta 21 de septiembre</td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT05" value="1" class="campo">&nbsp;</td><td class="celda">FDT05</td><td class="celda">Prevenci�n del riesgo qu�mico</td><td class="celda">10/05/2006</td><td class="celda">hasta 9 de mayo</td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT06" value="1" class="campo">&nbsp;</td><td class="celda">FDT06</td><td class="celda">Prevenci�n del riesgo qu�mico</td><td class="celda">11/09/2006</td><td class="celda">hasta 9 de mayo</td></tr>
</table>
<% end if %>
<br>&nbsp;

<div id="oculto3" style="overflow: auto; visibility: hidden; display: none"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Otros datos/opciones</td><td class=texto" align="right"><a onclick="mostrar('3')" style="cursor:hand">Mostrar campos [+]</a></td></tr></table></td></tr>
</table>
</div>

<div id="visible3" style="overflow: auto; visibility: visible; display: block"> 
<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td colspan="2"><table width="90%" align="center"><tr><td class="subtitulo" align="left" width="60%">Otros datos/opciones</td><td class=texto" align="right"><a onclick="ocultar('3')" style="cursor:hand">Ocultar campos [-]</a></td></tr> </table></td></tr>
<tr><td class="celda" align="left" colspan=2>Si te interesa adem�s recibir informaci�n peri�dica sobre medio ambiente y salud laboral, por favor <br>completa los siguientes campos tambi�n:</td></tr>
<tr><td class="celda" align="left" colspan=2>Seg�n la normativa vigente en protecci�n de datos personales, necesitamos su consentimiento expl�cito para poder utilizar los datos.</td></tr>
<tr>
	<td class=celdacentro align=right valign="middle">Quiero recibir informaci&oacute;n peri&oacute;dica sobre<br>las actividades y publicaciones de ECOinformas:</td>
	<td class=celdacentro align=left><% CALL valores("002","recibir_info_ecoinformas","2")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="middle">Me interesa recibir informaci&oacute;n peri&oacute;dica sobre<br>temas de medio ambiente y salud laboral, distribuida por ISTAS:</td>
	<td class=celdacentro align=left><% CALL valores("002","recibir_info_istas","2")%></td>
</tr>
<tr>
	<td class=celdacentro align=right valign="top">Si quieres anotar cualquier observaci�n junto al formulario:</td>
	<td class=celdacentro align=left><textarea name="observaciones" cols=50 rows=5 class="campo" OnKeyDown="return checkMaxLength(this, event,3000)" OnSelect="storeSelection(this)"></textarea></td>
</tr>
</table>
</div>

<br>&nbsp;

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td class=texto align=center>
   <input type="button" value="ENVIAR DATOS" name="comprobar" class="boton" onclick="enviar()">
</td></tr>
<tr><td class=texto align=center>
(*) Los datos que nos facilites ser�n incorporados a un fichero bajo titularidad de ISTAS. La finalidad del tratamiento de sus datos la constituye la posibilidad de difusi�n por correo electr�nico y ordinario de informaci�n y materiales de ECOinformas; la promoci�n de la salud laboral y la protecci�n del medio ambiente a trav�s de la remisi�n de informaci�n sobre los productos editoriales y actividades de ISTAS; auditor�a por parte de la Fundaci�n Biodiversidad que se compromete a su vez a cumplir la Ley Org�nica de Protecci�n de Datos de car�cter Personal (LOPD). Para m�s informaci�n: <a href="http://www.istas.net/ecoinformas/index.asp?idpagina=558" target="_blank">pol�tica de privacidad.</a>
</td></tr>
   
</table>
</form>

<!-- fin formulario -->				
				</div>
				<p>&nbsp;</p>
			</div>

			<map name="Map1" id="Map1">
            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
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
