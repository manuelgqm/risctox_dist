<!--#include file="eco_conexion.asp"-->
<HTML>

<head>
<title>Panel de control: visitas</title>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
</head>


<body bgcolor="#FEFEFE" topmargin="20" leftmargin="20">
<script src="valida_fecha.js"></script>
<%	
	fecha1 = EliminaInyeccionSQL(request("fecha1"))
	fecha2 = EliminaInyeccionSQL(request("fecha2"))
	if cstr(fecha1)="" then fecha1 = cdate(date()-1)
	if cstr(fecha2)="" then fecha2 = date()
	ip = EliminaInyeccionSQL(request("ip"))
	ordenador = EliminaInyeccionSQL(request("ordenador"))
	navegador = EliminaInyeccionSQL(request("navegador"))
	seccion = EliminaInyeccionSQL(request("seccion"))
	subpaginas = EliminaInyeccionSQL(request("subpaginas"))
	perfil = EliminaInyeccionSQL(request("perfil"))
	if cstr(ordenador)<>"" then textoordenador=" AND WEBISTAS_VISITAS.idgente="&clng(ordenador)
	if cstr(seccion)<>"" then 
		if subpaginas="si" then 
			textoseccion=" AND WEBISTAS_PAGINAS.numeracion LIKE '"&seccion&"%' "
		else
			textoseccion=" AND WEBISTAS_PAGINAS.numeracion='"&seccion&"' "
		end if
	end if
	agrupado = EliminaInyeccionSQL(request("agrupado"))
	if cstr(agrupado)="" then agrupado="dia"


	orden = "SELECT elegible_2007,count(idgente) as cuantos FROM ECOINFORMAS_GENTE GROUP BY elegible_2007 ORDER BY elegible_2007"
	Set objRecordset0 = Server.CreateObject ("ADODB.Recordset")
	Set objRecordset0 = OBJConnection.Execute(orden)
	if not objRecordset0.eof then
			num_sinasignar = objRecordset0("cuantos")
		else
			num_sinasignar = 0
	end if
	objRecordset0.movenext
	if not objRecordset0.eof then
		num_noelegibles = objRecordset0("cuantos")
	else
		num_noelegibles = 0
	end if
	objRecordset0.movenext
	if not objRecordset0.eof then
		num_elegibles = objRecordset0("cuantos")
	else
		num_elegibles = 0
	end if
	objRecordset0.movenext
	if not objRecordset0.eof then
		num_administradores = objRecordset0("cuantos")
	else
		num_administradores = 0
	end if
	num_todos = num_sinasignar+num_noelegibles+num_elegibles+num_administradores
	
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	
	if agrupado="no" then
		ordensql = "SELECT WEBISTAS_VISITAS.fecha, WEBISTAS_VISITAS.hora, WEBISTAS_VISITAS.ip, WEBISTAS_VISITAS.navegador, WEBISTAS_VISITAS.idgente, WEBISTAS_PAGINAS.titulo AS nombreseccion"
		ordensql = ordensql & " FROM WEBISTAS_VISITAS LEFT JOIN WEBISTAS_PAGINAS ON WEBISTAS_VISITAS.idpagina = WEBISTAS_PAGINAS.idpagina"
		if perfil<>"" then ordensql = ordensql & " LEFT JOIN ECOINFORMAS_GENTE ON WEBISTAS_VISITAS.idgente = ECOINFORMAS_GENTE.idgente"
		ordensql = ordensql & " WHERE WEBISTAS_PAGINAS.numeracion>'AI' AND WEBISTAS_PAGINAS.numeracion<'AJ' AND WEBISTAS_VISITAS.fecha<='"&fecha2&"' AND WEBISTAS_VISITAS.fecha>='"&fecha1&"' AND WEBISTAS_VISITAS.ip LIKE '%"&ip&"%' AND WEBISTAS_VISITAS.navegador LIKE '%"&navegador&"%'"&textoordenador&textoseccion
		if perfil="administradores" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=3"
		if perfil="elegibles" then ordensql = ordensql & " AND "
		if perfil="noelegibles" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=1"
		if perfil="sinasignar" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=0"
		ordensql = ordensql & " ORDER BY WEBISTAS_VISITAS.fecha DESC, WEBISTAS_VISITAS.hora DESC;"
	end if
	if agrupado="dia" then
		ordensql = "SELECT WEBISTAS_VISITAS.fecha, Count(WEBISTAS_VISITAS.fecha) AS visitaspordia"
		ordensql = ordensql & " FROM WEBISTAS_VISITAS LEFT JOIN WEBISTAS_PAGINAS ON WEBISTAS_VISITAS.idpagina = WEBISTAS_PAGINAS.idpagina"
		if perfil<>"" then ordensql = ordensql & " LEFT JOIN ECOINFORMAS_GENTE ON WEBISTAS_VISITAS.idgente = ECOINFORMAS_GENTE.idgente"
		ordensql = ordensql & " WHERE WEBISTAS_PAGINAS.numeracion>'AI' AND WEBISTAS_PAGINAS.numeracion<'AJ' AND WEBISTAS_VISITAS.ip LIKE '%"&ip&"%' AND WEBISTAS_VISITAS.navegador LIKE '%"&navegador&"%'"&textoordenador&textoseccion
		if perfil="administradores" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=3"
		if perfil="elegibles" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=2"
		if perfil="noelegibles" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=1"
		if perfil="sinasignar" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=0"
		ordensql = ordensql & " GROUP BY WEBISTAS_VISITAS.fecha"
		ordensql = ordensql & " HAVING WEBISTAS_VISITAS.fecha<='"&fecha2&"' AND WEBISTAS_VISITAS.fecha>='"&fecha1&"' "
		ordensql = ordensql & " ORDER BY WEBISTAS_VISITAS.fecha DESC;"
	end if
	if agrupado="ip" then
		ordensql = "SELECT WEBISTAS_VISITAS.ip, Count(WEBISTAS_VISITAS.ip) AS visitasporsesion"
		ordensql = ordensql & " FROM WEBISTAS_VISITAS LEFT JOIN WEBISTAS_PAGINAS ON WEBISTAS_VISITAS.idpagina = WEBISTAS_PAGINAS.idpagina"
		if perfil<>"" then ordensql = ordensql & " LEFT JOIN ECOINFORMAS_GENTE ON WEBISTAS_VISITAS.idgente = ECOINFORMAS_GENTE.idgente"
		ordensql = ordensql & " WHERE WEBISTAS_PAGINAS.numeracion>'AI' AND WEBISTAS_PAGINAS.numeracion<'AJ' AND WEBISTAS_VISITAS.fecha<='"&fecha2&"' AND WEBISTAS_VISITAS.fecha>='"&fecha1&"' AND WEBISTAS_VISITAS.navegador LIKE '%"&navegador&"%'"&textoordenador&textoseccion
		if perfil="administradores" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=3"
		if perfil="elegibles" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=2"
		if perfil="noelegibles" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=1"
		if perfil="sinasignar" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=0"
		ordensql = ordensql & " GROUP BY WEBISTAS_VISITAS.ip"
		ordensql = ordensql & " HAVING WEBISTAS_VISITAS.ip LIKE '%"&ip&"%'"
		ordensql = ordensql & " ORDER BY WEBISTAS_VISITAS.ip DESC;"
	end if	
	if agrupado="seccion" then
		ordensql = "SELECT WEBISTAS_PAGINAS.titulo AS nombreseccion,WEBISTAS_PAGINAS.tipo, WEBISTAS_PAGINAS.numeracion, Count(WEBISTAS_VISITAS.idpagina) AS visitasporapartado"
		ordensql = ordensql & " FROM WEBISTAS_VISITAS LEFT JOIN WEBISTAS_PAGINAS ON WEBISTAS_VISITAS.idpagina = WEBISTAS_PAGINAS.idpagina"
		if perfil<>"" then ordensql = ordensql & " LEFT JOIN ECOINFORMAS_GENTE ON WEBISTAS_VISITAS.idgente = ECOINFORMAS_GENTE.idgente"
		ordensql = ordensql & " WHERE WEBISTAS_PAGINAS.numeracion>'AI' AND WEBISTAS_PAGINAS.numeracion<'AJ' AND WEBISTAS_VISITAS.fecha<='"&fecha2&"' AND WEBISTAS_VISITAS.fecha>='"&fecha1&"' AND WEBISTAS_VISITAS.ip LIKE '%"&ip&"%' AND WEBISTAS_VISITAS.navegador LIKE '%"&navegador&"%'"&textoordenador
		if perfil="administradores" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=3"
		if perfil="elegibles" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=2"
		if perfil="noelegibles" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=1"
		if perfil="sinasignar" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=0"
		ordensql = ordensql & " GROUP BY WEBISTAS_PAGINAS.numeracion, WEBISTAS_PAGINAS.titulo, WEBISTAS_PAGINAS.tipo"
		if cstr(seccion)<>"" and subpaginas="si" then ordensql = ordensql & " HAVING WEBISTAS_PAGINAS.numeracion LIKE '"&seccion&"%' "
		if cstr(seccion)<>"" and subpaginas="no" then ordensql = ordensql & " HAVING WEBISTAS_PAGINAS.numeracion='"&seccion&"' "
		ordensql = ordensql & " ORDER BY WEBISTAS_PAGINAS.numeracion;"
	end if	
	if agrupado="ordenador" then
		ordensql = "SELECT WEBISTAS_VISITAS.idgente, Count(WEBISTAS_VISITAS.idgente) AS visitasporordenador"
		ordensql = ordensql & " FROM WEBISTAS_VISITAS LEFT JOIN WEBISTAS_PAGINAS ON WEBISTAS_VISITAS.idpagina = WEBISTAS_PAGINAS.idpagina"
		if perfil<>"" then ordensql = ordensql & " LEFT JOIN ECOINFORMAS_GENTE ON WEBISTAS_VISITAS.idgente = ECOINFORMAS_GENTE.idgente"
		ordensql = ordensql & " WHERE WEBISTAS_PAGINAS.numeracion>'AI' AND WEBISTAS_PAGINAS.numeracion<'AJ' AND WEBISTAS_VISITAS.fecha<='"&fecha2&"' AND WEBISTAS_VISITAS.fecha>='"&fecha1&"' AND WEBISTAS_VISITAS.navegador LIKE '%"&navegador&"%'"&textoordenador&textoseccion
		if perfil="administradores" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=3"
		if perfil="elegibles" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=2"
		if perfil="noelegibles" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=1"
		if perfil="sinasignar" then ordensql = ordensql & " AND ECOINFORMAS_GENTE.elegible_2007=0"
		ordensql = ordensql & " GROUP BY WEBISTAS_VISITAS.idgente"
		'if cstr(seccion)<>"" then ordensql = ordensql & " HAVING WEBISTAS_VISITAS.ip LIKE '%"&ip&"%'"
		ordensql = ordensql & " ORDER BY WEBISTAS_VISITAS.idgente;"
	end if
	
	'response.write ordensql
	'if 1=0 then
	objrecordset.Open ordensql,OBJConnection,adOpenKeyset
	numvisitas = objrecordset.recordcount

%>

<form name="formulario" action="eco_visitas.asp" METHOD="POST">
<table border="0" cellspacing="0" width="100%">
<tr><td class="cue_titulo"><b>ESTADÍSTICAS DE VISITAS A RIESGO QUÍMICO</b><br><br></td></tr>
<tr><td  class="cue_fuente">Elige las fechas en que quieres conocer las visitas, la IP (conexión), el navegador (IE o Netscape)<br>&nbsp;</td></tr>
<tr><td class="cue_fuente">
Desde:&nbsp;<input type="text" value="<%=fecha1%>" name="fecha1" size="10" maxlength="10" class="campo" ONBLUR="valida_fecha(fecha1);">&nbsp;&nbsp;
hasta:&nbsp;<input type="text" value="<%=fecha2%>" name="fecha2" size="10" maxlength="10" class="campo" ONBLUR="valida_fecha(fecha2);">&nbsp;&nbsp;
Dirección IP:&nbsp;<input type="text" value="<%=ip%>" name="ip" size="15" maxlength="15" class="campo">&nbsp;&nbsp;
Usuario/a núm:&nbsp;<input type="text" value="<%=ordenador%>" name="ordenador" size="3" maxlength="4" class="campo">&nbsp;&nbsp;
Navegador:&nbsp;<input type="text" value="<%=navegador%>" name="navegador" size="15" maxlength="15" class="campo">&nbsp;&nbsp;
Página:&nbsp;<select name="seccion" class="campo">
<option value="">- TODAS -</option>
<%
	Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
	ordensql2 = "SELECT idpagina,titulo,numeracion FROM WEBISTAS_PAGINAS WHERE WEBISTAS_PAGINAS.numeracion>'AI' AND WEBISTAS_PAGINAS.numeracion<'AJ' ORDER BY numeracion"
	objrecordset2.Open ordensql2,OBJConnection,adOpenKeyset
	do while not objrecordset2.eof
		if cstr(objrecordset2("numeracion"))=cstr(request("seccion")) then 
			textosel = "SELECTED"
		else
			textosel = ""
		end if
		titulo = ""
		titulo2 = titulo & " (" & objrecordset2("idpagina") & ")"
		numeracion = objrecordset2("numeracion")
		for i = 3 to len(numeracion)
			titulo = titulo & cstr(asc(mid(numeracion,i,1))-64) & "."
		next
		titulo = titulo & " " & left(objrecordset2("titulo"),60-2*len(numeracion))
%>
<option value="<%=objrecordset2("numeracion")%>" <%=textosel%>><%=titulo%></option>
<%  objrecordset2.movenext
	loop
%>
</select>&nbsp;&nbsp;
incluir subpáginas:&nbsp;<select name="subpaginas" class="campo">
<option value="si" <% if subpaginas="si" then response.write "selected"%>>sí</option>
<option value="no" <% if subpaginas="no" then response.write "selected"%> >no</option>
</select>&nbsp;&nbsp;&nbsp;

Agrupadas:&nbsp;<select name="agrupado" class="campo">
<option value="no">no</option>
<%		if agrupado="dia" then 
			textosel = "SELECTED"
		else
			textosel = ""
		end if
%>
<option value="dia" <%=textosel%>>por día</option>
<%		if agrupado="seccion" then 
			textosel = "SELECTED"
		else
			textosel = ""
		end if
%>
<option value="seccion" <%=textosel%>>por página</option>
<%		if agrupado="ordenador" then 
			textosel = "SELECTED"
		else
			textosel = ""
		end if
%>
<option value="ordenador" <%=textosel%>>por usuario</option>
<%		if agrupado="ip" then 
			textosel = "SELECTED"
		else
			textosel = ""
		end if
%>
<option value="ip" <%=textosel%>>por conexión</option>
</select>&nbsp;&nbsp;
<input type="submit" value="VISITAS" class="boton"></td></tr>
</table>
</form>
<p class="cue_fuente" align="center">El número de resultados registrados con estos parámetros es de <b><%=numvisitas%></b></p>

<table class="cue_fuente" border="0" width="95%" align="center">
<%if agrupado="no" then%>
<tr>
<td class="cue_celda" style="text-align:left"><b>PÁGINA</b></td>
<td class="cue_celda" style="text-align:left"><b>FECHA</b></td>
<td class="cue_celda" style="text-align:left"><b>HORA</b></td>
<td class="cue_celda" style="text-align:left"><b>DIRECCIÓN IP</b></td>
<td class="cue_celda" style="text-align:left"><b>NAVEGADOR</b></td>
</tr>
<% 
i=1
do while not objRecordset.eof 
  if i=1 then 
  	colorcelda = "#DDDDDD"
  else
  	colorcelda = "#EEEEEE"
  end if
  i=-1*i
%>
<tr>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("nombreseccion")%></td>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("fecha")%></td>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("hora")%></td>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("ip")%></td>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("navegador")%></td>
</tr>
<% objrecordset.movenext
   loop
end if%>

<%if agrupado="dia" then%>
<tr>
<td class="cue_celda" style="text-align:left"><b>FECHA</b></td>
<td class="cue_celda" style="text-align:left"><b>PÁGINAS</b></td>
</tr>
<% 
i=1
do while not objRecordset.eof 
  if i=1 then 
  	colorcelda = "#DDDDDD"
  else
  	colorcelda = "#EEEEEE"
  end if
  i=-1*i
  visitas = visitas+objRecordset("visitaspordia")
%>
<tr>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("fecha")%></td>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("visitaspordia")%></td>
</tr>
<% objrecordset.movenext
   loop %>
<% if numvisitas>0 then %>
<tr><td class="cue_fuente" colspan="2" align="right"><b>Total=<%=visitas%>. Media=<%=formatnumber(visitas/numvisitas,2,0,0,-1)%>&nbsp; páginas vistas por día</b></td></tr>
<%end if%>
<%end if%>

<%if agrupado="ip" then%>
<tr>
<td class="cue_celda" style="text-align:left"><b>CONEXIÓN</b></td>
<td class="cue_celda" style="text-align:left"><b>PÁGINAS</b></td>
</tr>
<% 
i=1
do while not objRecordset.eof 
  if i=1 then 
  	colorcelda = "#DDDDDD"
  else
  	colorcelda = "#EEEEEE"
  end if
  i=-1*i
  visitas = visitas+objRecordset("visitasporsesion")
%>
<tr>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("ip")%></td>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("visitasporsesion")%></td>
</tr>
<% objrecordset.movenext
   loop %>
<% if numvisitas>0 then %>
<tr><td class="cue_fuente" colspan="2" align="right"><b>Total=<%=visitas%>. Media=<%=formatnumber(visitas/numvisitas,2,0,0,-1)%>&nbsp; páginas vistas por conexión</b></td></tr>
<%end if%>
<%end if%>

<%if agrupado="seccion" then%>
<tr>
<td class="cue_celda" style="text-align:left"><b>PÁGINAS</b></td>
<td class="cue_celda" style="text-align:left"><b>VISITAS</b></td>
</tr>
<% 
i=1
do while not objRecordset.eof 
  if i=1 then 
  	colorcelda = "#DDDDDD"
  else
  	colorcelda = "#EEEEEE"
  end if
  i=-1*i
  visitas = visitas+objRecordset("visitasporapartado")
%>
<tr>
<td class="cue_fuente" bgcolor="<%=colorcelda%>">
<% if objRecordset("tipo")=7 then
 	response.write "<b>"&objRecordset("nombreseccion")&"</b>"
   else
 	response.write objRecordset("nombreseccion")
   end if
 %>

</td>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("visitasporapartado")%></td>
</tr>
<% objrecordset.movenext
   loop %>
<% if numvisitas>0 then %>
<tr><td class="cue_fuente" colspan="2" align="right"><b>Total=<%=visitas%>. Media=<%=formatnumber(visitas/numvisitas,2,0,0,-1)%>&nbsp; visitas a cada página</b></td></tr>
<%end if%>
<%end if%>

<%if agrupado="ordenador" then%>
<tr>
<td class="cue_celda" style="text-align:left"><b>NÚM. USUARIO/A</b></td>
<td class="cue_celda" style="text-align:left"><b>PÁGINAS</b></td>
</tr>
<% 
i=1
do while not objRecordset.eof 
  if i=1 then 
  	colorcelda = "#DDDDDD"
  else
  	colorcelda = "#EEEEEE"
  end if
  i=-1*i
  visitas = visitas+objRecordset("visitasporordenador")

  idgente = objRecordset("idgente")
  nombre = ""
  ordensql2 = "SELECT nombre,apellidos FROM ECOINFORMAS_GENTE WHERE idgente="&objRecordset("idgente")
  set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
  objrecordset2.Open ordensql2,OBJConnection,adOpenKeyset
  if not objrecordset2.eof then nombre = ". "&objrecordset2("nombre")&" "&objrecordset2("apellidos")
%>
<tr>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=idgente&nombre%></td>
<td class="cue_fuente" bgcolor="<%=colorcelda%>"><%=objRecordset("visitasporordenador")%></td>
</tr>
<% objrecordset.movenext
   loop %>
<% if numvisitas>0 then %>
<tr><td class="cue_fuente" colspan="2" align="right"><b>Total=<%=visitas%>. Media=<%=formatnumber(visitas/numvisitas,2,0,0,-1)%>&nbsp; páginas vistas por usuario/a</b></td></tr>
<%end if%>
<%end if%>


</table>
</body>

</HTML>
<% 'end if %>
<!--


<% response.write ordensql %>


-->