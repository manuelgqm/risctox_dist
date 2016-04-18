<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

  if session("id_ecogente")="" then response.redirect "formulario2.asp"
	

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
<title>ECOinformas: usuario identificado</title>
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
	c01 = document.formulario.seg_social.value;
  c02 = document.formulario.dni.value;
  c03 = document.formulario.puesto.value;
  c04 = document.formulario.contrato.value;
  c05 = document.formulario.estudios.value;
  c06 = document.formulario.direccion.value;
  c07 = document.formulario.localidad.value;
  c08 = document.formulario.provincia.value;
  c09 = document.formulario.cp.value;
  c10 = document.formulario.movil.value;
  c11 = document.formulario.fax.value;
  c12 = document.formulario.empresa.value;
  c13 = document.formulario.cif.value;
  c14 = document.formulario.sector.value;
  c15 = document.formulario.emp_direccion.value;
  c16 = document.formulario.emp_localidad.value;
  c17 = document.formulario.emp_cp.value;
  c18 = document.formulario.emp_telefono.value;
  c19 = document.formulario.emp_fax.value;
  c20 = document.formulario.emp_movil.value;
  c21 = document.formulario.emp_web.value;
  c22 = document.formulario.emp_email.value;
  c23 = document.formulario.relacion_ma.value;
  c24 = document.formulario.rlt.value;
  c25 = document.formulario.emp_tipo.value;
  c26 = document.formulario.emp_facturacion.value;
  c27 = document.formulario.emp_provincia.value;
  
	falta = '';
	if (c01=='') { falta = falta+'seguridad social'+'\n'; }
  if (c02=='') { falta = falta+'dni'+'\n'; }
  if (c03=='') { falta = falta+'puesto'+'\n'; }
  if (c04=='') { falta = falta+'tipo contrato'+'\n'; }
  if (c05=='') { falta = falta+'estudios'+'\n'; }
  if (c06=='') { falta = falta+'dirección'+'\n'; }
  if (c07=='') { falta = falta+'localidad'+'\n'; }
  if (c08=='') { falta = falta+'provincia'+'\n'; }
  if (c09=='') { falta = falta+'código postal'+'\n'; }
  if (c10=='') { falta = falta+'movil'+'\n'; }
  if (c11=='') { falta = falta+'fax'+'\n'; }
  if (c12=='') { falta = falta+'empresa'+'\n'; }
  if (c13=='') { falta = falta+'cif'+'\n'; }
  if (c14=='') { falta = falta+'sector'+'\n'; }
  if (c15=='') { falta = falta+'dirección de la empresa'+'\n'; }
  if (c16=='') { falta = falta+'localidad de la empresa'+'\n'; }
  if (c17=='') { falta = falta+'código postal de la empresa'+'\n'; }
  if (c18=='') { falta = falta+'teléfono empresa'+'\n'; }
  if (c19=='') { falta = falta+'fax empresa'+'\n'; }
  if (c20=='') { falta = falta+'móvil empresa'+'\n'; }
  if (c21=='') { falta = falta+'web empresa'+'\n'; }
  if (c22=='') { falta = falta+'email empresa'+'\n'; }
  if (c23=='') { falta = falta+'relación con el medioambiente'+'\n'; }
  if (c24=='') { falta = falta+'RLT'+'\n'; }
  if (c25=='') { falta = falta+'tipo de empresa'+'\n'; }
  if (c26=='') { falta = falta+'facturación empresa'+'\n'; }
	if (c27=='') { falta = falta+'provincia empresa'+'\n'; }

	if (falta!='')
	{  alert ('Falta por rellenar:\n\n'+falta) }
	else
	{  document.formulario.submit(); }
}

// -->
</SCRIPT>
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
<%            			sql = "SELECT * FROM ECOINFORMAS_GENTE WHERE idgente="&session("id_ecogente")
			   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	        set objRecordset = OBJConnection.Execute(sql)
		   	   	        usuario = objRecordset("nombre")&" "&objRecordset("apellidos")
		   	   	        usuario_sexo = "o"
		   	   	        if objRecordset("sexo")=75 then usuario_sexo = "a"
		   	   	        FDT01_2007 = objRecordset("FDT01_2007")
		   	   	        FDT02_2007 = objRecordset("FDT02_2007")
		   	   	        FDT03_2007 = objRecordset("FDT03_2007")
		   	   	        FDT04_2007 = objRecordset("FDT04_2007")
		   	   	        FDT05_2007 = objRecordset("FDT05_2007")
		   	   	        FDT06_2007 = objRecordset("FDT06_2007")
		   	   	        seg_social = objRecordset("seg_social")
		   	   	        dni = objRecordset("dni")
		   	   	        puesto = objRecordset("puesto")
		   	   	        contrato = objRecordset("contrato")
		   	   	        estudios = objRecordset("estudios")
		   	   	        direccion = objRecordset("direccion")
		   	   	        localidad = objRecordset("localidad")
		   	   	        provincia = objRecordset("provincia")
		   	   	        cp = objRecordset("cp")
		   	   	        movil = objRecordset("movil")
		   	   	        fax = objRecordset("fax")
		   	   	        empresa = objRecordset("empresa")
		   	   	        cif = objRecordset("cif")
		   	   	        'razon_social = objRecordset("razon_social")
		   	   	        sector = objRecordset("sector")
		   	   	        emp_direccion = objRecordset("emp_direccion")
		   	   	        emp_localidad = objRecordset("emp_localidad")
		   	   	        emp_cp = objRecordset("emp_cp")
		   	   	        emp_telefono = objRecordset("emp_telefono")
		   	   	        emp_fax = objRecordset("emp_fax")
		   	   	        emp_movil = objRecordset("emp_movil")
		   	   	        emp_web = objRecordset("emp_web")
		   	   	        emp_email = objRecordset("emp_email")
		   	   	        relacion_ma = objRecordset("relacion_ma")
		   	   	        rlt = objRecordset("rlt")
		   	   	        emp_tipo = objRecordset("emp_tipo")
		   	   	        emp_facturacion = objRecordset("emp_facturacion")
		   	   	        emp_provincia = objRecordset("emp_provincia")
		   	   	        
		   	   	        
%>
            			<tr><td align="right">Usuari<%=usuario_sexo%> identificad<%=usuario_sexo%>:&nbsp;<%=usuario%></td></tr>
          		</table>
			</div>
       			<% end if %>
			
			<div id="texto">
			
				
				<div class="texto">
<!--- formulario -->				
				
<form method="POST" name="formulario" action="formulario_identificado_grabar.asp">

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">
<tr><td class="titulo" align="center" colspan=5>Solicitud de cursos ECOinformas</td></tr> 
<%if 1=0 then %>
<tr><td class="texto" align="justify" colspan=5>Si quieres cambiar algún dato o darte de baja escribe un email a <a href=mailto:datospersonales@istas.net>datospersonales@istas.net</a>.</td></tr>
<% end if %>
<tr><td class="texto" align="justify" colspan=5>Si perteneces a alguno de los <a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=557" target="_blank">colectivos a los que va dirigido este proyecto</a> y quieres realizar alguno de los cursos gratuitos on-line, selecciona el o los cursos que te interesan y rellena los siguientes campos adicionales:</td></tr>
<tr><td>&nbsp;</td><td><strong>Código</strong></td><td><strong>Acción formativa</strong></td><td><strong>Fecha inicio</strong></td><td><strong>Plazo inscripción</strong></td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT01_2007" value="1" class="campo" <% if FDT01_2007=1 then response.write "checked"%>>&nbsp;</td><td class="celda">FDT01</td><td class="celda">Curso on-line sobre "Introducción a los Sistemas de Gestión Medioambiental"</td><td class="celda">26/04/2007</td><td class="celda">18/04/2007</td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT02_2007" value="1" class="campo" <% if FDT02_2007=1 then response.write "checked"%>>&nbsp;</td><td class="celda">FDT02</td><td class="celda">Curso on-line sobre "Introducción a los Sistemas de Gestión Medioambiental"</td><td class="celda">03/05/2007</td><td class="celda">25/04/2007</td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT05_2007" value="1" class="campo" <% if FDT05_2007=1 then response.write "checked"%>>&nbsp;</td><td class="celda">FDT05</td><td class="celda">Curso on-line sobre "Energías Renovables, Medio Ambiente y Empleo"</td><td class="celda">04/05/2007</td><td class="celda">26/04/2007</td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT03_2007" value="1" class="campo" <% if FDT03_2007=1 then response.write "checked"%>>&nbsp;</td><td class="celda">FDT03</td><td class="celda">Curso on-line Básico de Riesgo Químico</td><td class="celda">09/05/2007</td><td class="celda">30/04/2007</td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT06_2007" value="1" class="campo" <% if FDT06_2007=1 then response.write "checked"%>>&nbsp;</td><td class="celda">FDT06</td><td class="celda">Curso on-line sobre "Energías Renovables, Medio Ambiente y Empleo"</td><td class="celda">11/05/2007</td><td class="celda">03/05/2007</td></tr>
<tr><td class="celda" align="right"><input type="checkbox" name="FDT04_2007" value="1" class="campo" <% if FDT04_2007=1 then response.write "checked"%>>&nbsp;</td><td class="celda">FDT04</td><td class="celda">Curso on-line Básico de Riesgo Químico</td><td class="celda">16/05/2007</td><td class="celda">08/05/2007</td></tr>
</table>

<table cellspacing=0 border=0 cellpadding=5 class=tabla width="95%" align="center">

<tr><td colspan="2" align="center" class="subtitulo">Datos personales adicionales</td></tr>
<% if isnull(dni) or dni="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">DNI/NIE/Otro:</td>
	<td class=celdacentro align=left><input type="text" name="dni" size="20" maxlength="20" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="dni" value="<%=dni%>">
<% end if %>

<% if isnull(seg_social) or seg_social="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Núm. Seguridad Social:</td>
	<td class=celdacentro align=left><input type="text" name="seg_social" size="20" maxlength="20" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="seg_social" value="<%=seg_social%>">
<% end if %>

<% if isnull(estudios) or estudios="" or estudios=0 then %>
<tr>
	<td class=celdacentro align=right valign="middle">Datos académicos:</td>
	<td class=celdacentro align=left><% CALL valores("007","estudios","1")%></td>
</tr>
<% else %>
<input type="hidden" name="estudios" value="<%=estudios%>">
<% end if %>

<% if isnull(direccion) or direccion="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Domicilio:</td>
	<td class=celdacentro align=left><input type="text" name="direccion" size="50" maxlength="200" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="direccion" value="<%=direccion%>">
<% end if %>

<% if isnull(localidad) or localidad="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Localidad:</td>
	<td class=celdacentro align=left><input type="text" name="localidad" size="50" maxlength="200" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="localidad" value="<%=localidad%>">
<% end if %>

<% if isnull(provincia) or provincia="" or provincia=0 then %>
<tr>
	<td class=celdacentro align=right valign="middle">Provincia:</td>
	<td class=celdacentro align=left><% CALL valores("013","provincia","1")%></td>
</tr>
<% else %>
<input type="hidden" name="provincia" value="<%=provincia%>">
<% end if %>

<% if isnull(cp) or cp="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Código postal:</td>
	<td class=celdacentro align=left><input type="text" name="cp" size="5" maxlength="20" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="cp" value="<%=cp%>">
<% end if %>

<% if isnull(movil) or movil="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Telefono móvil:</td>
	<td class=celdacentro align=left><input type="text" name="movil" size="20" maxlength="100" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="movil" value="<%=movil%>">
<% end if %>

<% if isnull(fax) or fax="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Fax:</td>
	<td class=celdacentro align=left><input type="text" name="fax" size="20" maxlength="100" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="fax" value="<%=fax%>">
<% end if %>

<tr>
	<td colspan="2" class=celdacentro align=right valign="middle">&nbsp;</td>
</tr>

<% if puesto="" or contrato="" or isnull(puesto) or isnull(contrato) or isnull(relacion_ma) or relacion_ma="" or isnull(rlt) or rlt="" then %>
<tr>
	<td class=subtitulo align=center valign="middle" colspan="2">Datos laborales</td>
</tr>
<% end if %>

<% if isnull(puesto) or puesto="" or puesto=0 then %>
<tr>
	<td class=celdacentro align=right valign="middle">Puesto que desempeñas:</td>
	<td class=celdacentro align=left><% CALL valores("005","puesto","1")%></td>
</tr>
<% else %>
<input type="hidden" name="puesto" value="<%=puesto%>">
<% end if %>

<% if isnull(contrato) or contrato="" or contrato=0 then %>
<tr>
	<td class=celdacentro align=right valign="middle">Tipo de contratación:</td>
	<td class=celdacentro align=left><% CALL valores("006","contrato","1")%></td>
</tr>
<% else %>
<input type="hidden" name="contrato" value="<%=contrato%>">
<% end if %>

<% if isnull(relacion_ma) or relacion_ma="" or relacion_ma=0 then %>
<tr>
	<td class=celdacentro align=right valign="middle">¿Está tu actividad relacionada con el medioambiente?:</td>
	<td class=celdacentro align=left><% CALL valores("002","relacion_ma","2")%></td>
</tr>
<% else %>
<input type="hidden" name="relacion_ma" value="<%=relacion_ma%>">
<% end if %>

<% if isnull(rlt) or rlt="" or rlt=0 then %>
<tr>
	<td class=celdacentro align=right valign="middle">¿Eres representante legal de los trabajadores (RLT)?:</td>
	<td class=celdacentro align=left><% CALL valores("002","rlt","2")%></td>
</tr>
<% else %>
<input type="hidden" name="rlt" value="<%=rlt%>">
<% end if %>

<tr>
	<td colspan="2" class=celdacentro align=right valign="middle">&nbsp;</td>
</tr>

<% 'if cif="" or emp_direccion="" or emp_localidad="" or emp_provincia="" or emp_cp="" or emp_telefono="" or emp_movil="" or emp_fax="" or emp_email="" or emp_web="" or emp_tipo="" or sector="" or puesto="" or emp_facturacion="" then %>
<tr>
	<td class=subtitulo align=center valign="middle" colspan="2">Datos de la empresa</td>
</tr>
<% 'end if %>

<% if isnull(empresa) or empresa="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Empresa:</td>
	<td class=celdacentro align=left><input type="text" name="empresa" size="50" maxlength="200" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="empresa" value="<%=empresa%>">
<% end if %>

<% if isnull(cif) or cif="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">CIF:</td>
	<td class=celdacentro align=left><input type="text" name="cif" size="20" maxlength="20" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="cif" value="<%=cif%>">
<% end if %>

<% if isnull(emp_direccion) or emp_direccion="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Dirección:</td>
	<td class=celdacentro align=left><input type="text" name="emp_direccion" size="50" maxlength="200" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="emp_direccion" value="<%=emp_direccion%>">
<% end if %>

<% if isnull(emp_localidad) or emp_localidad="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Localidad:</td>
	<td class=celdacentro align=left><input type="text" name="emp_localidad" size="50" maxlength="200" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="emp_localidad" value="<%=emp_localidad%>">
<% end if %>

<% if isnull(emp_provincia) or emp_provincia="" or emp_provincia=0 then %>
<tr>
	<td class=celdacentro align=right valign="middle">Provincia:</td>
	<td class=celdacentro align=left><% CALL valores("013","emp_provincia","1")%></td>
</tr>
<% else %>
<input type="hidden" name="emp_provincia" value="<%=emp_provincia%>">
<% end if %>

<% if isnull(emp_cp) or emp_cp="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Código postal:</td>
	<td class=celdacentro align=left><input type="text" name="emp_cp" size="5" maxlength="20" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="emp_cp" value="<%=emp_cp%>">
<% end if %>

<% if isnull(emp_telefono) or emp_telefono="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Telefono:</td>
	<td class=celdacentro align=left><input type="text" name="emp_telefono" size="20" maxlength="100" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="emp_telefono" value="<%=emp_telefono%>">
<% end if %>

<% if isnull(emp_movil) or emp_movil="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Teléfono móvil:</td>
	<td class=celdacentro align=left><input type="text" name="emp_movil" size="20" maxlength="100" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="emp_movil" value="<%=emp_movil%>">
<% end if %>

<% if isnull(emp_fax) or emp_fax="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Fax:</td>
	<td class=celdacentro align=left><input type="text" name="emp_fax" size="20" maxlength="100" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="emp_fax" value="<%=emp_fax%>">
<% end if %>

<% if isnull(emp_email) or emp_email="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Email:</td>
	<td class=celdacentro align=left><input type="text" name="emp_email" size="50" maxlength="200" class="campo"></td>
</tr>
<% else %>
<input type="hidden" name="emp_email" value="<%=emp_email%>">
<% end if %>

<% if isnull(emp_web) or emp_web="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Página web:</td>
	<td class=celdacentro align=left><input type="text" name="emp_web" size="50" maxlength="200" class="campo" value="http://"></td>
</tr>
<% else %>
<input type="hidden" name="emp_web" value="<%=emp_web%>">
<% end if %>

<% if isnull(emp_tipo) or emp_tipo="" or emp_tipo=0 then %>
<tr>
	<td class=celdacentro align=right valign="middle">Tipo de empresa:</td>
	<td class=celdacentro align=left><% CALL valores("017","emp_tipo","1")%></td>
</tr>
<% else %>
<input type="hidden" name="emp_tipo" value="<%=emp_tipo%>">
<% end if %>

<% if isnull(sector) or sector="" or sector=0 then %>
<tr>
	<td class=celdacentro align=right valign="middle">Familia profesional:</td>
	<td class=celdacentro align=left><% CALL valores("008","sector","1")%></td>
</tr>
<% else %>
<input type="hidden" name="sector" value="<%=sector%>">
<% end if %>

<% if isnull(emp_facturacion) or emp_facturacion="" then %>
<tr>
	<td class=celdacentro align=right valign="middle">Facturación aproximada:</td>
	<td class=celdacentro align=left><input type="text" name="emp_facturacion" size="20" maxlength="50" class="campo">&nbsp;euros</td>
</tr>
<% else %>
<input type="hidden" name="emp_facturacion" value="<%=emp_facturacion%>">
<% end if %>

<tr>
	<td colspan="2" class=celdacentro align=right valign="middle">&nbsp;</td>
</tr>

<tr><td class="texto" align="center" colspan=5>&nbsp;</td></tr> 
<tr><td class="texto" align="center" colspan=5><b>(recuerda que todos los campos son obligatorios)</b></td></tr> 
<tr><td class="texto" align="center" colspan=5>
   <input type="button" value="ENVIAR DATOS" name="comprobar" class="boton" onclick="enviar()">
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
