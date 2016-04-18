<!--#include file="eco_conexion.asp"-->
<%

	numero = EliminaInyeccionSQL(request("numero"))
	buscar = EliminaInyeccionSQL(request("buscar"))
	fecini = EliminaInyeccionSQL(request("fecini"))
	fecfin = EliminaInyeccionSQL(request("fecfin"))
	autor = EliminaInyeccionSQL(request("autor"))
	
	sql = "SELECT idpagina,titulo,fecha,hora,fecha_modificacion,numeracion,tipo,visible FROM WEBISTAS_PAGINAS WHERE numeracion LIKE 'AI%' AND len(numeracion)>1 AND "
	
	if numero<>"" then 
		sql = sql & " (idpagina="&numero&") "
	else
		if buscar<>"" then sql = sql & " (titulo LIKE '%"&buscar&"%' OR pagina LIKE '%"&buscar&"%') AND "
		if fecini<>"" then sql = sql & " (fecha_modificacion>='"&fecini&"') AND "
		if fecfin<>"" then sql = sql & " (fecha_modificacion<='"&fecfin&"') AND "
		if cstr(autor)<>"0" and cstr(autor)<> "" then sql = sql & " (autor='"&autor&"') AND "
		sql = sql & " 1=1 ORDER BY numeracion"
	end if
	
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	
%>
<html>

<head>
<title>Buscador de páginas</title>
<base target="_self">
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
</head>
<body bgcolor="#FEFEFE" topmargin="10" leftmargin="15">
<table>
<% if objrecordset.eof then %>
<tr><td class="cue_fuente">No hay ninguna página con esos valores de búsqueda</td></tr>
<% else %>
<% do while not objrecordset.eof 
	  numeracion = objrecordset("numeracion")
	  situacion = asc(mid(numeracion,2,1))-64
	  for i = 3 to len(numeracion)
	   situacion = situacion&"."&(asc(mid(numeracion,i,1))-64)
	  next
%>
<tr><td class="cue_fuente">
<% 	  if objrecordset("tipo")=0 then 
		estilo = "azulindice"
	  else
	  	if objrecordset("visible")<>1 then 
	  		estilo = "rojoindice"
	  	else
	  		if objrecordset("tipo")=1 then estilo = "negroindice"
	  		if objrecordset("tipo")=2 then estilo = "azulindice"
	  	end if
	  end if
	  if isnull(objrecordset("fecha_modificacion")) then
	  	fecha_modificacion = "?"
  	  else
  		fecha_modificacion = objrecordset("fecha_modificacion")
  	  end if
%>	  	
<a href="eco_editarpagina.asp?id=<%=objrecordset("idpagina")%>"><font class="<%=estilo%>" style="text-decoration:none"><%=situacion&". "&objrecordset("titulo")%></a></td></tr>
<tr><td class="peque">Última actualización:&nbsp;<%=fecha_modificacion%>. Página <%=objrecordset("idpagina")%></td></tr>
<tr><td class="cue_fuente">&nbsp;</td></tr>
<% objrecordset.movenext
   loop %>
<% end if %>
</table>
</body>
</html>
