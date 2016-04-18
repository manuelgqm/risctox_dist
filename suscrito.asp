<!--#include file="EliminaInyeccionSQL.asp"-->
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"


	dim campo(71,4)
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
	campo(66,0)="clave" 
	campo(67,0)="contra" 
	campo(68,0)="confirmado_web" 
	campo(69,0)="confirmado_cursos" 
	campo(70,0)="confirmado_materiales" 
	campo(71,0)="fec_hor" 
	
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
	campo(63,3)="SEP01"
	campo(64,3)="SEP02"
	campo(65,3)="SEP03"
	campo(66,3)="clave" 
	campo(67,3)="contraseña" 
	campo(68,3)="confirmado web" 
	campo(69,3)="confirmado cursos" 
	campo(70,3)="confirmado materiales" 
	campo(71,3)="fecha y hora de la suscripción" 

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
	campo(39,4)="0"
	campo(40,4)="0"
	campo(41,4)="0"
	campo(42,4)="0"
	campo(43,4)="0"
	campo(44,4)="0"
	campo(45,4)="0"
	campo(46,4)="1"
	campo(47,4)="0"
	campo(48,4)="0"
	campo(49,4)="1"
	campo(50,4)="1"
	campo(51,4)="1"
	campo(52,4)="1"
	campo(53,4)="1"
	campo(54,4)="1"
	campo(55,4)="1"
	campo(56,4)="1"
	campo(57,4)="1"
	campo(58,4)="1"
	campo(59,4)="1"
	campo(60,4)="1"
	campo(61,4)="1"
	campo(62,4)="1"
	campo(63,4)="1"
	campo(64,4)="1"
	campo(65,4)="1"
	campo(66,4)="1"
	campo(67,4)="1"
	campo(68,4)="1"
	campo(69,4)="1"
	campo(70,4)="1"
	campo(71,4)="1"

	
	if EliminaInyeccionSQL(request("id"))<>"" and EliminaInyeccionSQL(request("id"))<>"0" then
		orden = "SELECT * FROM ECOINFORMAS_GENTE WHERE idgente="&EliminaInyeccionSQL(request("id"))
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		Set objRecordset = OBJConnection.Execute(orden)

		for i=1 to 71
			if campo(i,4)=1 then campo(i,2) = objrecordset(campo(i,0))
		next	
		clave = campo(66,2)
		contra = ucase(campo(67,2))
		'if campo(63,2)="" then
		'	'clave = "ECO" & right("000"&cstr(objrecordset("idgente")),4)
		'	clave = "ECO" & cstr(objrecordset("idgente"))
		'else
		'	clave = campo(63,2)
		'end if
        	
		'if campo(64,2)="" then
		'	apellidos = split(campo(2,2)," ")
		'	contra = ucase(apellidos(0))
		'else
		'	contra = ucase(campo(64,2))
		'end if				
	else
		for i=1 to 71
			campo(i,2) = ""
		next			
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


FUNCTION valores (vgrupo,vname,tipo,vselec)

orden2 = "SELECT * FROM ECOINFORMAS_VALORES WHERE grupo='"&vgrupo&"' ORDER BY valor,desc1"
Set dSQL2 = Server.CreateObject ("ADODB.Recordset")
dSQL2.Open orden2,objConnection,adOpenKeyset
if tipo="1" then 				'--- tipo 1 = desplegable
%>	<select name=<%=vname%> class="campo">
	<option value="">- Selecciona de la lista -</option><%
	if not(DSQL2.bof and DSQL2.eof) then
		dSQL2.movefirst
		DO while not dSQL2.eof
		if cstr(vselec)=cstr(dSQL2("valor")) then
			sele = "selected"
		else
			sele = ""
		end if
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
	<input type="radio" name="<%=vname%>" class="campo" value="<%=dSQL2("valor")%>" <%if cstr(vselec)=cstr(dSQL2("valor")) then response.write "checked"%>>&nbsp;<%=dSQL2("desc1")%>&nbsp;&nbsp;
	<%      dSQL2.movenext
	        loop
	end if
end if
dSQL2.close        
END FUNCTION

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Description" content="ISTAS Instituto Sindical de Trabajo, Ambiente y Salud - Medio Ambiente">
<meta name="Author" content="XiP multimèdia">
<title>suscritos ECOinformas</title>
<link rel="stylesheet" type="text/css" href="/intra2004/intranet.css">
</head>

<body bgcolor="#FFFFFF" link="#666699" vlink="#6A4F9A" alink="#6A4F9A" class="cuerpo">

<form name="formulario" action="suscrito_modificar.asp">
<input type="hidden" name="id" value="<%=EliminaInyeccionSQL(request("id"))%>">
<table cellspacing=0 border=0 cellpadding=5 width="95%" align="center">
<tr><td colspan="2" class="campocamuflado" align="center"><b>Datos de acceso</b></td></tr>
<tr>
	<td class=negro align=right valign="middle">Fecha y hora de la suscripción:</td>
	<td class=negro align=left><%=campo(71,2)%></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Elegible (acceso web):</td>
	<td class=negro align=left>
	<select name="confirmado_web" class="campo">
		<option value="0" <%if campo(68,2)="0" then response.write "selected"%>>sin asignar</option>
		<option value="1" <%if campo(68,2)="1" then response.write "selected"%>>no</option>
		<option value="2" <%if campo(68,2)="2" then response.write "selected"%>>sí</option>
		<option value="3" <%if campo(69,2)="3" then response.write "selected"%>>administrador</option>
	</select></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Acceso Cursos:</td>
	<td class=negro align=left>
	<select name="confirmado_cursos" class="campo">
		<option value="0" <%if campo(69,2)="0" then response.write "selected"%>>sin asignar</option>
		<option value="1" <%if campo(69,2)="1" then response.write "selected"%>>pendiente</option>
		<option value="3" <%if campo(69,2)="3" then response.write "selected"%>>no</option>
		<option value="2" <%if campo(69,2)="2" then response.write "selected"%>>sí</option>
	</select></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Acceso Materiales:</td>
	<td class=negro align=left>
	<select name="confirmado_materiales" class="campo">
		<option value="0" <%if campo(70,2)="0" then response.write "selected"%>>sin asignar</option>
		<option value="1" <%if campo(70,2)="1" then response.write "selected"%>>pendiente</option>
		<option value="3" <%if campo(70,2)="3" then response.write "selected"%>>no</option>
		<option value="2" <%if campo(70,2)="2" then response.write "selected"%>>sí</option>
	</select></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Clave:</td>
	<td class=negro align=left><input type="text" name="clave" size="10" maxlength="10" class="campo" value="<%=clave%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Contraseña:</td>
	<td class=negro align=left><input type="text" name="contra" size="10" maxlength="10" class="campo" value="<%=contra%>"></td>
</tr>
	

<tr><td colspan="2" class="campocamuflado" align="center"><b>Datos personales</b></td></tr>
<tr>
	<td class=negro align=right valign="middle">Nombre:</td>
	<td class=negro align=left><input type="text" name="nombre" size="50" maxlength="200" class="campo" value="<%=campo(1,2)%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Apellidos:</td>
	<td class=negro align=left><input type="text" name="apellidos" size="50" maxlength="200" class="campo" value="<%=campo(2,2)%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Fecha nacimiento:</td>
	<td class=negro align=left><input type="text" name="fec_nac" size="11" maxlength="10" class="campo" OnBlur='valida_fecha(this)' value="<%=campo(3,2)%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Sexo:</td>
	<td class=negro align=left><% CALL valores("001","sexo","2",campo(4,2))%></td>
</tr>
<% if 1=0 then %>
<tr>
	<td class=negro align=right valign="middle">Núm. Seguridad Social:</td>
	<td class=negro align=left><input type="text" name="seg_social" size="50" maxlength="50" class="campo" value="<%=campo(5,2)%>"></td>
</tr>
<% end if %>
<tr>
	<td class=negro align=right valign="middle">Minusvalía reconocida:</td>
	<td class=negro align=left><% CALL valores("002","minusvalia","2",campo(6,2))%></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Eres inmigrante:</td>
	<td class=negro align=left><% CALL valores("002","inmigrante","2",campo(7,2))%></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Eres trabajador de baja cualificación:</td>
	<td class=negro align=left><% CALL valores("002","cualificacion","2",campo(8,2))%></td>
</tr>
<% if 1=0 then %>
<tr>
	<td class=negro align=right valign="middle">DNI/NIE:</td>
	<td class=negro align=left><input type="text" name="dni" size="10" maxlength="9" class="campo" OnBlur='valida_dni(this)' value="<%=campo(9,2)%>">&nbsp;(escribe sólo los números, la letra sale automáticamente)</td>
</tr>
<% end if %>
<tr>
	<td class=negro align=right valign="middle">Condición laboral:</td>
	<td class=negro align=left><% CALL valores("003","cond_laboral","1",campo(10,2))%></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Tamaño de tu empresa:</td>
	<td class=negro align=left><% CALL valores("004","tam_empresa","1",campo(11,2))%></td>
</tr>
<% if 1=0 then %>
<tr>
	<td class=negro align=right valign="middle">Puesto que desempeñas:</td>
	<td class=negro align=left><% CALL valores("005","puesto","1",campo(12,2))%></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Tipo de contrataci&oacute;n:</td>
	<td class=negro align=left><% CALL valores("006","contrato","1",campo(13,2))%></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Datos acad&eacute;micos:</td>
	<td class=negro align=left><% CALL valores("007","estudios","1",campo(14,2))%></td>
</tr>
<% end if %>
<tr>
	<td class=negro align=right valign="middle">Direcci&oacute;n (residencia habitual):</td>
	<td class=negro align=left><input type="text" name="direccion" size="50" maxlength="200" class="campo" value="<%=campo(15,2)%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Localidad:</td>
	<td class=negro align=left><input type="text" name="localidad" size="50" maxlength="200" class="campo" value="<%=campo(16,2)%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Provincia:</td>
	<td class=negro align=left><% CALL valores("013","provincia","1",campo(17,2))%></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">C&oacute;digo postal:</td>
	<td class=negro align=left><input type="text" name="cp" size="5" maxlength="5" class="campo" value="<%=campo(18,2)%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Tel&eacute;fono:</td>
	<td class=negro align=left><input type="text" name="telefono" size="50" maxlength="50" class="campo" value="<%=campo(19,2)%>"></td>
</tr>
<% if 1=0 then %>
<tr>
	<td class=negro align=right valign="middle">M&oacute;vil:</td>
	<td class=negro align=left><input type="text" name="movil" size="50" maxlength="50" class="campo" value="<%=campo(20,2)%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Fax:</td>
	<td class=negro align=left><input type="text" name="fax" size="50" maxlength="50" class="campo" value="<%=campo(21,2)%>"></td>
</tr>
<% end if %>
<tr>
	<td class=negro align=right valign="middle">Email:</td>
	<td class=negro align=left><input type="text" name="email" size="50" maxlength="200" class="campo" ONBLUR='valida_mail(this)' value="<%=campo(22,2)%>"></td>
</tr>
<tr><td colspan="2" class=campocamuflado align="center"><b>Datos de la empresa en que trabaja</b></td></tr>
<tr>
	<td class=negro align=right valign="middle">Nombre empresa:</td>
	<td class=negro align=left><input type="text" name="empresa" size="50" maxlength="200" class="campo" value="<%=campo(23,2)%>"></td>
</tr>
<% if 1=0 then %>
<tr>
	<td class=negro align=right valign="middle">CIF:</td>
	<td class=negro align=left><input type="text" name="cif" size="10" maxlength="10" class="campo" value="<%=campo(24,2)%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Raz&oacute;n social:</td>
	<td class=negro align=left><input type="text" name="razon_social" size="50" maxlength="200" class="campo" value="<%=campo(25,2)%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Sector:</td>
	<td class=negro align=left><% CALL valores("008","sector","1",campo(26,2))%></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Direcci&oacute;n:</td>
	<td class=negro align=left><input type="text" name="emp_direccion" size="50" maxlength="200" class="campo" value="<%=campo(27,2)%>"></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Localidad:</td>
	<td class=negro align=left><input type="text" name="emp_localidad" size="50" maxlength="200" class="campo" value="<%=campo(28,2)%>"></td>
</tr>
<% end if %>
<tr>
	<td class=negro align=right valign="middle">Provincia:</td>
	<td class=negro align=left><% CALL valores("013","emp_provincia","1",campo(29,2))%></td>
</tr>


<tr><td colspan="2" class="campocamuflado" align="center"><b>Otros datos/opciones</b></td></tr>
<tr>
	<td class=negro align=right valign="middle">Quiero recibir informaci&oacute;n peri&oacute;dica sobre<br>las actividades y publicaciones de ECOinformas:</td>
	<td class=negro align=left><% CALL valores("002","recibir_info_ecoinformas","2",campo(36,2))%></td>
</tr>
<tr>
	<td class=negro align=right valign="middle">Me interesa recibir informaci&oacute;n peri&oacute;dica sobre<br>temas de medio ambiente y salud laboral, distribuida por ISTAS:</td>
	<td class=negro align=left><% CALL valores("002","recibir_info_istas","2",campo(37,2))%></td>
</tr>
<tr>
	<td class=negro align=right valign="top">Observaciones:</td>
	<td class=negro align=left><textarea name="observaciones" cols=50 rows=5 class="campo" OnKeyDown="return checkMaxLength(this, event,3000)" OnSelect="storeSelection(this)"><%=campo(38,2)%></textarea></td>
</tr>
</table>


</form>
</body>
</html>
