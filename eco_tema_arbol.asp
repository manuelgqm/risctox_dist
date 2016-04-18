<!--#include file="eco_conexion3.asp"-->
<%
	id_tema = EliminaInyeccionSQL(request("idtema"))
	if id_tema<>"" then
	  sql = "SELECT * FROM WEBISTAS_PAGINAS where idpagina="&id_tema
	else
	  sql = "SELECT * FROM WEBISTAS_PAGINAS where numeracion='AI'"
	end if
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	if not objRecordset.eof then
		numeracion_tema = objRecordset("numeracion")
		nivel_tema = len(numeracion_tema)
		id_tema = objrecordset("idpagina")
	else
		numeracion_tema = "AI"
		nivel_tema = len(numeracion_tema)
		id_tema = 1
		id_tema = 542
	end if

	mos1 = EliminaInyeccionSQL(request("mos1"))
	mos2 = EliminaInyeccionSQL(request("mos2"))
	mos3 = EliminaInyeccionSQL(request("mos3"))
	mos4 = EliminaInyeccionSQL(request("mos4"))
	mos5 = EliminaInyeccionSQL(request("mos5"))
	mos6 = EliminaInyeccionSQL(request("mos6"))
	ord  = EliminaInyeccionSQL(request("ord"))
	fil1 = EliminaInyeccionSQL(request("fil1"))
	fil2 = EliminaInyeccionSQL(request("fil2"))
	fil3 = EliminaInyeccionSQL(request("fil3"))
	fil4 = EliminaInyeccionSQL(request("fil4"))
	fil5 = EliminaInyeccionSQL(request("fil5"))
	fil6 = EliminaInyeccionSQL(request("fil6"))

	if mos1="" and mos2="" and mos3="" and mos4="" and mos5="" and mos6="" then
		mos1="1"
		mos2="1"
		mos3="1"
	end if

	param = "mos1="&mos1&"&mos2="&mos2&"&mos3="&mos3&"&mos4="&mos4&"&mos5="&mos5&"&mos6="&mos6&"&ord="&ord&"&fil1="&fil1&"&fil2="&fil2&"&fil3="&fil3&"&fil4="&fil4&"&fil5="&fil5&"&fil6="&fil6

	function contenga(codigo)
		texto_sql = "("
		for n=1 to len(codigo)-1
			texto_sql = texto_sql & " ascii(substring(numeracion,"&cstr(n)&",1))=" & asc(mid(codigo,n,1)) & " AND "
		next
		texto_sql = texto_sql & " ascii(substring(numeracion,"&cstr(len(codigo))&",1))=" & asc(mid(codigo,len(codigo),1)) & ")"
		contenga = texto_sql
		'response.write texto_sql & "<br>"
	end function

	sql = "SELECT WEBISTAS_PAGINAS.idpagina,WEBISTAS_PAGINAS.titulo,WEBISTAS_PAGINAS.tipo,WEBISTAS_PAGINAS.visible,WEBISTAS_PAGINAS.numeracion,WEBISTAS_PAGINAS.fecha,WEBISTAS_PAGINAS.fecha_modificacion FROM WEBISTAS_PAGINAS "
	sql = sql & "WHERE "
	if fil1="" then
		'sql = sql & "((numeracion like 'AI%' AND len(numeracion)<=2) "
		sql = sql & "(( numeracion ='B' OR (" & contenga("AI") & " AND len(numeracion)<=3)) "								' mostrar los niveles principales
		for k=3 to nivel_tema
	  	  'sql = sql & "OR (numeracion like '"& mid(numeracion_tema,1,i)&"%' AND len(numeracion)="&cstr(i+1)&") "
	  	  sql = sql & "OR (" & contenga(mid(numeracion_tema,1,k)) & " AND len(numeracion)="&cstr(k+1)&") "						' mostrar los padres
		next
		'if nivel_tema>1 then sql = sql & "OR (numeracion like '"&mid(numeracion_tema,1,nivel_tema-1)&"%' AND len(numeracion)="&cstr(nivel_tema)&") "
		if nivel_tema>2 then sql = sql & "OR (" & contenga(mid(numeracion_tema,1,nivel_tema-1)) & " AND len(numeracion)="&cstr(nivel_tema)&") "		' mostrar los hermanos
		'sql = sql & "OR (numeracion like '"&numeracion_tema&"%' AND len(numeracion)="&cstr(nivel_tema+1)&")) "
		sql = sql & "OR (" & contenga(numeracion_tema) & " AND len(numeracion)="&cstr(nivel_tema+1)&")) "						' mostrar los hijos
	else
		if fil1="AI" then
		  sql = sql & " (numeracion>'AI') "
		else
		  sql2 = "SELECT numeracion FROM WEBISTAS_PAGINAS WHERE idpagina="&fil1
		  set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		  objRecordset2.Open sql2,objConnection,adOpenKeyset
		  rama_selec = objRecordset2("numeracion")
		  'sql = sql & " (numeracion like '"& rama_selec&"%') "
		  sql = sql & " ("&contenga(rama_selec)&") "
		end if
	end if
	if fil2<>"" and fil3<>"" then
		sql = sql &" and fecha>='"&fil2&"' and fecha<='"&fil3&"' "
	end if
	if fil4<>"" and fil5<>"" then
		sql = sql &" and fecha_modificacion>='"&fil4&"' and fecha_modificacion<='"&fil5&"' "
	end if
	if fil6<>"" then
		sql = sql &" and WEBISTAS_PAGINAS.titulo LIKE '%"&fil6&"%' "
	end if

	sql = sql & "ORDER BY "
	if ord="" or ord="ord_num" then
		for w=1 to 9
		  sql = sql & "ascii(substring(numeracion,"&cstr(w)&",1)),"
		next
		sql = sql & "ascii(substring(numeracion,10,1));"
	end if
	if ord="ord_nom" then
		sql = sql & "WEBISTAS_PAGINAS.titulo;"
	end if
	if ord="ord_fal" then
		sql = sql & "WEBISTAS_PAGINAS.fecha;"
	end if
	if ord="ord_fmo" then
		sql = sql & "WEBISTAS_PAGINAS.fecha_modificacion;"
	end if
	if ord="ord_aut" then
		sql = sql & "WEBISTAS_PAGINAS.autor;"
	end if

	'response.write sql
	'if 1=0 then

	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,objConnection,adOpenKeyset
	numero_de_temas = objrecordset.recordcount

%>

<html>
<head>
<title>P&aacute;ginas en &aacute;rbol</title>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
<script>

function brillo(que)
{	que.style.color = "#ff6600";  }

function mate(que)
{	que.style.color = "#000000";  }

function ayuda(idpagina)
{	window.status = "Página número "+idpagina; }


function cambiatema(idtema)
{	parent.parent.frames.contenido.location.href='eco_editarpagina.asp?id='+idtema;
	location.href='eco_tema_arbol.asp?<%=param%>&idtema='+idtema; }

</script>
</head>

<body bgcolor="#FEFEFE" topmargin="10" leftmargin="15">
<form name="botones">
<table width="90%"><tr>
<td align="right">
<% if 1=0 then %>
<img src="nuevo.gif" class="boton" alt="Editar rama" onClick="window.open('tema_editar.asp?<%=param%>&idtema=<%=id_tema%>','editartema','width=400,height=400,resizable=yes,scrollbars=yes,status=yes')">&nbsp;
<% end if %>
<img src="ver.gif" class="boton" alt="Opciones del listado de ramas" onClick="window.open('eco_tema_opcioneslistado.asp?<%=param%>&idtema=<%=id_tema%>','vertemas','scrollbars=yes,width=400,height=400')">&nbsp;
<img src="imprimir.gif" class="boton" alt="Imprimir listado de ramas" onClick="print()">&nbsp;
</td></tr>
</table>
</form>
<table border="0">
<tr><td class=negroindice><b>ESPACIO WEB Riesgo químico</b>&nbsp;<font class=gris>(ramas abiertas:&nbsp;<%=numero_de_temas%>)</font></td></tr>
<%
	j = 0
	Do while not objrecordset.eof
	  j = j+1
	  idpagina = objrecordset("idpagina")
	  numeracion = objrecordset("numeracion")
	  nivel = len(numeracion)
	  tipo = objrecordset("tipo")
	  visible = objrecordset("visible")
	  titulo = replace(objrecordset("titulo"),"'","&#39;")
	  imagen = "<img src=flecha_ver.gif align=absmiddle>"
	  if tipo<4 then imagen = "<img src=ico_pagina_eco.gif align=absmiddle>"
	  if tipo=4 then imagen = "<img src=ico_ficha.gif align=absmiddle>"
	  if tipo=5 then imagen = "<img src=ico_carp_fichas.gif align=absmiddle>"
	  if tipo=6 then imagen = "<img src=ico_carp_enlaces.gif align=absmiddle>"
	  if tipo=7 or tipo=8 then imagen = "<img src=ico_pagina_restringida.gif align=absmiddle>"
	  if numeracion="B" then imagen = "<img src=papelera.gif align=absmiddle>"

	  if visible=1 then
	  	colorfuente="negroindice"
	  else
	  	colorfuente="rojoindice"
	  end if
	  if cstr(idpagina)=cstr(id_tema) then
	  	fondo = ";background=#FFFFFF"
	  else
	  	fondo = ""
	  end if
	  'if tipo=7 then fondo = fondo & "; border:2 solid #FF0000"

	  if len(numeracion)>2 then
	  	situacion = asc(mid(numeracion,3,1))-64
	  	for i2 = 4 to nivel
	   	  situacion = situacion & "." & (asc(mid(numeracion,i2,1))-64)
	  	next
	  	situacion = situacion & ".&nbsp;"
	  else
	  	situacion = ""
	  end if

	  response.write "<tr><td class=negroindice nowrap><table class=negroindice cellpadding=0 cellspacing=0 border=0><tr>"
	  if mos1="1" then response.write "<td width="&cstr((nivel-1)*15)&" nowrap>&nbsp;</td>"
	  response.write "<td valign=top>"
	  if mos3="1" then response.write imagen&"&nbsp;"
	  if mos1="1" then response.write situacion
	  response.write "</td>"
	  if mos2="1" then response.write "<td class="&colorfuente&" valign=bottom nowrap><font style='text-decoration:none; cursor:hand; "&fondo&"' onmouseover='ayuda("&idpagina&")' onmouseout='ayuda("&idpagina&")' onclick=cambiatema("&idpagina&")>"&(titulo)&"</font>"
	  if mos4="1" then response.write "<br><font style=color:#666666>Fecha alta: "&objrecordset("fecha")&"</font>"
	  if mos5="1" then response.write "<br><font style=color:#666666>Fecha modificación: "&objrecordset("fecha_modificacion")&"</font>"
	  if mos2="1" then response.write "</td>"
	  response.write "</tr></table></td></tr>"&chr(13)

	  objrecordset.movenext
	loop

%>
</table>
</body>

</html>
<% 'end if %>