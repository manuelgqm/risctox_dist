<!--#include file="eco_conexion2.asp"-->
<%
		
	id = EliminaInyeccionSQL(request("id"))
	if id="" then
		id = 1186
	else
		id = clng(id)
	end if
	
	'Registro con los datos de la página 
	sql = "SELECT titulo,numeracion,pagina,fecha,hora,autor,tipo,visible,destino, titulo_eng, pagina_eng FROM WEBISTAS_PAGINAS WHERE idpagina="&id
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	titulo = objRecordset("titulo")	
	titulo_eng = objRecordset("titulo_eng")	
	numeracion = objRecordSet("numeracion")
	numeracion_completa = numeracion
	pagina = objRecordset("pagina")
	pagina_eng = objRecordset("pagina_eng")
	if isnull(pagina) then pagina=" "
	fechapagina = objrecordset("fecha")
	horaficha = objrecordset("hora")
	propietarioficha = objrecordset("autor")
	tipo = objrecordset("tipo")
	visible = objrecordset("visible")
	destino = objrecordset("destino")
	
	objRecordset.close
	
	sql = "SELECT numeracion FROM WEBISTAS_PAGINAS WHERE numeracion LIKE '"&numeracion&"%' AND len(numeracion)="&cstr(len(numeracion)+1)&" ORDER BY numeracion DESC"
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,OBJConnection,adOpenKeyset
	if objrecordset.eof then
		hijomayor = 0
	else
		letra_hijomayor = objrecordset("numeracion")
		hijomayor = asc(mid(letra_hijomayor,(len(numeracion)+1),1))-64
	end if


	'sql = "SELECT * FROM ISTAS_GENTE WHERE idgente="&clng(propietarioficha)
	'Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	'set objRecordset = OBJConnection.Execute(sql)
	'if not objrecordset.eof then propietarioficha = objrecordset("nombre")&" "&objrecordset("apellidos")
	
	sql = "SELECT count(idvisita) as cuantas_visitas FROM WEBISTAS_VISITAS WHERE idpagina="&clng(id)
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	if not objrecordset.eof then cuantas_visitas = objrecordset("cuantas_visitas")
	
	
	FUNCTION vistaprevia(texto)
		if vartype(texto) = 1 then texto = ""
		texto = replace(texto,chr(13),"<br>")
		texto = replace(texto,"'","&#39;")
		texto = replace(texto,"<v1>","<img src=../../ecoinformas/web/imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v2>","&nbsp;&nbsp;<img src=../../ecoinformas/web/imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v3>","&nbsp;&nbsp;&nbsp;&nbsp;<img src=../../ecoinformas/web/imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v4>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=../../ecoinformas/web/imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v5>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=../../ecoinformas/web/imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v6>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=../../ecoinformas/web/imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<pag=","<a href=index.asp?idpagina=")
		texto = replace(texto,"</pag>","</a>")
		texto = replace(texto,"<e=","<a target=_blank href=abreenlace.asp?idenlace=")
		texto = replace(texto,"</e>","</a>")
		texto = replace(texto,"<t>","<font class=titulo3>")
		texto = replace(texto,"</t>","</font>")
		texto = replace(texto,"<st>","<font class=subtitulo3>")
		texto = replace(texto,"</st>","</font>")
		texto = replace(texto,"<pd>","<table width=95% align=center cellpadding=10 cellspacing=0 class=tabla><tr><td>")
		texto = replace(texto,"</pd>","</td><td valign=top align=center><img src=../../ecoinformas/web/pd.gif></td></tr></table>")
		vistaprevia = texto
		
	END FUNCTION
	
	FUNCTION chora(expresion)
	 x = instr(expresion," ")
	 chora = right(expresion,len(expresion)-x)
	END FUNCTION

%>	


<html>

<head>
<title>Edición de la página <%=id%></title>
<base target="_self">

</head>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
<!-- <link rel="stylesheet" type="text/css" href="ecoinformas.css"> -->
<body bgcolor="#FEFEFE" topmargin="20" leftmargin="20" onLoad="document.formulario.titulo.focus()">
<script LANGUAGE="JScript">
<!--
// 
    ua=navigator.userAgent; v=navigator.appVersion.substring(0,1);
    if ((ua.lastIndexOf("MSIE")!=-1) && (v!='1') && (v!='2') && (v!='3')) { document.body.onmouseover=Brillo; document.body.onmouseout=Mate }
    function Brillo() { src = event.toElement; if (src.tagName == "A") { src.antes = src.style.color; src.style.color ="#F09450" } }
    function Mate() { src=event.fromElement; if (src.tagName == "A") { src.style.color = src.antes } }



function comprueba_pagina()
{
	
	//if (document.formulario.titulo.value != "")
		document.formulario.submit();
	//else
	//	alert('Introduce el título de la página');	
}

function formato(codigo)
{
	if (document.selection.createRange().text!='')
	{ document.selection.createRange().text = '<'+codigo+'>'+document.selection.createRange().text+'</'+codigo+'>' }
	else
	{ alert('Hay que seleccionar un texto para poder aplicarle formato') }
}

function formatopagina(codigo)
{
	if (document.selection.createRange().text!='')
	{ if (codigo!='') 
	   { document.selection.createRange().text = '<pag='+codigo+'>'+document.selection.createRange().text+'</pag>' }
	   else
	   { alert('Elige la página a la que quieres vincular el texto seleccionado'); }
	}
	else
	{ alert('Hay que seleccionar un texto para poder vincular una página') }
}

function formatoficha(codigo)
{
	if (document.selection.createRange().text!='')
	{ if (codigo!='') 
	   { document.selection.createRange().text = '<f='+codigo+'>'+document.selection.createRange().text+'</f>' }
	   else
	   { alert('Elige la ficha a la que quieres vincular el texto seleccionado'); }
	}
	else
	{ alert('Hay que seleccionar un texto para poder vincular una ficha') }
}

function formatoenlace(codigo)
{
	if (document.selection.createRange().text!='')
	{ if (codigo!='') 
	   { document.selection.createRange().text = '<e='+codigo+'>'+document.selection.createRange().text+'</e>' }
	   else
	   { alert('Elige el enlace al que quieres vincular el texto seleccionado'); }
	}
	else
	{ alert('Hay que seleccionar un texto para poder vincular un enlace') }
}


function formatoenlacerestringido
(codigo)
{
	if (document.selection.createRange().text!='')
	{ if (codigo!='') 
	   { document.selection.createRange().text = '<er='+codigo+'>'+document.selection.createRange().text+'</e>' }
	   else
	   { alert('Elige el enlace al que quieres vincular de forma restringida el texto seleccionado'); }
	}
	else
	{ alert('Hay que seleccionar un texto para poder vincular un enlace') }
}
function formatocarpetaenlaces(codigo)
{
	if (document.selection.createRange().text!='')
	{ if (codigo!='') 
	   { document.selection.createRange().text = '<c='+codigo+'>'+document.selection.createRange().text+'</c>' }
	   else
	   { alert('Elige la carpeta de enlaces a la que quieres vincular el texto seleccionado'); }
	}
	else
	{ alert('Hay que seleccionar un texto para poder vincular una carpeta de enlaces') }
}

function mover()
{
	fichero = 'eco_desplazar.asp?id=<%=id%>'
	for (i=1 ; i < 11; i++)
	{ fichero = fichero+'&c'+i+'=';
	  eval('fichero = fichero+document.formulario.clasif'+i+'.value');
	}
	location.href = fichero;
}

function incluirfichero(nombre,descripcion)
{
	indice = document.formulario.ficheros.value;
	nombre = document.formulario.ficheros.options[indice].text;
	descripcion = document.formulario.descripciones.options[indice].text;
	document.formulario.contenido.value = document.formulario.contenido.value + '<br><img src=fichero.gif border=0 align=texttop alt='+descripcion+'> <a href=ftp/'+nombre+' target=_blank>'+nombre+'</a>'
}

function incluirtabla(textohtml)
{
	document.formulario.contenido.value = document.formulario.contenido.value + textohtml + ' ';
}

//-->
</script>
<form name="formulario" ACTION="eco_modificarpagina.asp" METHOD="POST">
<!--------------------------------------------------------------------CLASIFICACIÓN--> 
<table><tr><td class="cue_celda"><b>Clasificación de la página <%=id%> (ha recibido <%=cuantas_visitas%> visitas):</b></td></tr></table>
<% for i=1 to 10
	valor = ""
	if len(numeracion)>=i then 
		valor = cstr(asc(mid(numeracion,i,1))-64)
		if i>1 then numeracioncompleta = numeracioncompleta&valor&"."
	end if
%>
<%	if i<3 then %>
<input type="hidden" name="clasif<%=i%>" value="<%=valor%>" size="2" class="campo" maxlength="2">
<%	else %>
<input type="text" name="clasif<%=i%>" value="<%=valor%>" size="2" class="campo" maxlength="2">.&nbsp;
<%	end if %>
<% next %>
&nbsp;<input type="button" class="boton" value="DESPLAZAR" onClick="mover()">
<br>
<%	titulocompleto = ""
	for i=2 to len(numeracion)
		sql = "SELECT titulo,numeracion,idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='" & mid(numeracion,1,i) & "'"
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		if i<>2 then titulocompleto = titulocompleto & "&nbsp;&gt;&nbsp;" 
		titulocompleto = titulocompleto & "<a href=editarpagina.asp?id="&objrecordset("idpagina")&">"&objrecordset("titulo")&"</a>"
	next 
	if mid(numeracion,1,1)="B" then response.write "<img src=papelera.gif align=absmiddle>&nbsp;"
	response.write "<font class=negroindice>"&titulocompleto&"<font>" %>

<br>

<!--------------------------------------------------------------------NUEVA SUBPÁGINA--> 
<p align="center"><input type="button" value="CREAR NUEVA PÁGINA <%=numeracioncompleta&cstr(hijomayor+1)%>" class="boton" onClick="location.href='eco_subpagina.asp?id=<%=id%>'">&nbsp;&nbsp;
<!--------------------------------------------------------------------AYUDA-->
<input type="button" value="?" class="boton" onClick="window.open('ayuda_codigos.asp','','width=400,height=300,resizable=yes,scrollbars=yes')">&nbsp;&nbsp;
<%if 1=0 then %>
<!--------------------------------------------------------------------ABRIR EN VENTANA NUEVA-->
<input type="button" value="ABRIR EN VENTANA NUEVA" class="boton" onClick="window.open('eco_editarpagina.asp?id=<%=id%>','ventana','width=600,height=480,status=yes,scrollbars=yes,resizable=yes')">
<% end if %>
</p>

<!--------------------------------------------------------------------TÍTULO--> 
<table><tr><td class="cue_celda"><b>Título:</b></td></tr></table>
<input type="hidden" name="id" value="<%=id%>" size="3" class="campo">
<input type="text" name="titulo" value="<%=titulo%>" size="100" class="campo" maxlength="255">
<table><tr><td class="cue_celda"><b>Título (en inglés):</b></td></tr></table>
<input type="text" name="titulo_eng" value="<%=titulo_eng%>" size="100" class="campo" maxlength="255">
<br><br>

<!--------------------------------------------------------------------TIPO--> 
<table width="90%"><tr><td class="cue_celda"><b>Tipo:</b></td>
<td class="cue_fuente">
<select class="campo" name="tipo">
<option value="2" class="azulindice" <%if tipo=2 then response.write "selected"%>>Página normal</option>
<option value="1" class="azulindice" <%if tipo=1 then response.write "selected"%>>Página HTML (no aplica los códigos)</option>
<option value="7" class="azulindice" <%if tipo=7 then response.write "selected"%>>Página normal restringida</option>
<option value="8" class="azulindice" <%if tipo=8 then response.write "selected"%>>Página HTML restringida</option>
<option value="3" class="azulindice" <%if tipo=3 then response.write "selected"%>>Otra página ya existente</option>
<option value="6" class="azulindice" <%if tipo=6 then response.write "selected"%>>Carpeta de enlaces de la bv</option>
</select>&nbsp;
<% if 1=0 then %>
Destino:&nbsp;<input type="text" name="destino" value="<%=destino%>" size="4" class="campo" maxlength="5">
<% end if %>
</td>
<td class="cue_fuente">
&nbsp;&nbsp;&nbsp;Visible&nbsp;<select class="campo" name="visible">
<option value="1" <%if visible=1 then response.write "selected"%>>S?/option>
<option value="2"  class="rojoindice" <%if visible<>1 then response.write "selected"%>>No</option>
</select>
</td>
</tr></table>
<br><br>
<!--------------------------------------------------------------------CONTENIDO--> 
<table bgcolor="#e3e3e3" width="620"><tr><td class="cue_celda"><b>Contenido:</b></td>
<td align="right">
<img src="negrita.gif" onClick="formato('b')" alt="negrita" style="cursor:hand">&nbsp;
<img src="cursiva.gif" onClick="formato('i')" alt="cursiva" style="cursor:hand">&nbsp;
<img src="subrayado.gif" onClick="formato('u')" alt="subrayada" style="cursor:hand">&nbsp;
<img src="titulo.gif" onClick="formato('t')" alt="título" style="cursor:hand">&nbsp;
<img src="subtit.gif" onClick="formato('st')" alt="subtítulo" style="cursor:hand">&nbsp;
<img src="parrafo.gif" onClick="formato('pd')" alt="parrafo destacado" style="cursor:hand">&nbsp;
<img src="ico_maq_imagen.gif" onClick="window.open('../../recursos_gestion/recursos.asp?formacion=si','recursos','width=1024,height=700,resizable=yes,scrollbars=yes,status=yes')" alt="incluir recursos multimedia" style="cursor:hand">&nbsp;
</td>
</tr></table>
<textarea rows="25" name="contenido" cols="100" class="campo"><%=pagina%>
</textarea><br>
<table bgcolor="#e3e3e3" width="620"><tr><td class="cue_celda"><b>Contenido (en inglés):</b></td>
<td align="right">
<img src="negrita.gif" onClick="formato('b')" alt="negrita" style="cursor:hand">&nbsp;
<img src="cursiva.gif" onClick="formato('i')" alt="cursiva" style="cursor:hand">&nbsp;
<img src="subrayado.gif" onClick="formato('u')" alt="subrayada" style="cursor:hand">&nbsp;
<img src="titulo.gif" onClick="formato('t')" alt="título" style="cursor:hand">&nbsp;
<img src="subtit.gif" onClick="formato('st')" alt="subtítulo" style="cursor:hand">&nbsp;
<img src="parrafo.gif" onClick="formato('pd')" alt="parrafo destacado" style="cursor:hand">&nbsp;
<img src="ico_maq_imagen.gif" onClick="window.open('../../recursos_gestion/recursos.asp?formacion=si','recursos','width=1024,height=700,resizable=yes,scrollbars=yes,status=yes')" alt="incluir recursos multimedia" style="cursor:hand">&nbsp;
</td>
</tr></table>
<textarea rows="25" name="contenido_eng" cols="100" class="campo"><%=pagina_eng%>
</textarea><br>
<table cellpadding="0" cellspacing="2">
<tr>
<td class="cue_fuente">RECURSO DE LA BIBLIOTECA VIRTUAL (BV):&nbsp;</td><td class="cue_fuente"><input type="text" class="campo" name="enlaces" size=4>&nbsp;</td><td class="cue_fuente">
<input type="button" value="VER" class="boton" onClick="window.open('listado_enlaces.asp','listado_enlaces','width=300,height=300,resizable=yes,scrollbars=yes,menubar=yes')">&nbsp;
<input type="button" value="VINCULAR" class="boton" onClick="formatoenlace(formulario.enlaces.value)">&nbsp;&lt;e=XXX&gt;
</td></tr>
<tr>
<td class="cue_fuente">RECURSO RESTRINGIDO DE LA BIBLIOTECA VIRTUAL:&nbsp;</td><td class="cue_fuente"><input type="text" class="campo" name="enlacesrestringidos" size=4>&nbsp;</td><td class="cue_fuente">
<input type="button" value="VER" class="boton" onClick="window.open('listado_enlacesrestringidos.asp','listado_enlacesrestringidos','width=300,height=300,resizable=yes,scrollbars=yes,menubar=yes')">&nbsp;
<input type="button" value="VINCULAR" class="boton" onClick="formatoenlacerestringido(formulario.enlacesrestringidos.value)">&nbsp;&lt;er=XXX&gt;
</td></tr>
<tr>
<td class="cue_fuente">CARPETA DE RECURSOS DE LA BV:&nbsp;</td><td class="cue_fuente"><input type="text" class="campo" name="carpetaenlaces" size=4>&nbsp;</td><td class="cue_fuente">
<input type="button" value="VER" class="boton" onClick="window.open('listado_carpetaenlaces.asp','listado_carpetaenlaces','width=300,height=300,resizable=yes,scrollbars=yes,menubar=yes')">&nbsp;
<input type="button" value="VINCULAR" class="boton" onClick="formatocarpetaenlaces(formulario.carpetaenlaces.value)">&nbsp;&lt;c=XXX&gt;
</td></tr>
<tr>
<td class="cue_fuente">PÁGINAS ECOinformas:&nbsp;</td><td class="cue_fuente"><input type="text" class="campo" name="listadopaginas" size=4>&nbsp;</td><td class="cue_fuente">
<input type="button" value="VER" class="boton" onClick="window.open('eco_listado_paginas.asp','listado_paginas','width=300,height=300,resizable=yes,scrollbars=yes,menubar=yes,status=yes')">&nbsp;
<input type="button" value="VINCULAR" class="boton" onClick="formatopagina(formulario.listadopaginas.value)">&nbsp;&lt;pag=XXX&gt;
</td></tr></table>
<br><br>

<table align="center"><tr><td><input type="button" value="GUARDAR CAMBIOS" class="boton" onClick="comprueba_pagina()"></td></tr></table>
<br>

<!--------------------------------------------------------------------VISTA PREVIA--> 
<table><tr><td class="cue_celda"><b>Vista previa:</b></td></tr></table>
<table width="550" class="cajacerrada" bgcolor="#FFFFFF" cellpadding="10">
<tr><td width="550" class="cue_fuente">
<%=vistaprevia(pagina)%>
</td></tr>
</table>
<br>

<table><tr><td class="cue_celda"><b>Vista previa (en inglés):</b></td></tr></table>
<table width="550" class="cajacerrada" bgcolor="#FFFFFF" cellpadding="10">
<tr><td width="550" class="cue_fuente">
<%=vistaprevia(pagina_eng)%>
</td></tr>
</table>
<br>

<!--------------------------------------------------------------------ÚLTIMA ACTUALIZACIÓN--> 
<table><tr><td class="cue_fuente">Última actualización:&nbsp;<%=fechapagina%> a las <%=chora(horaficha)%></td>
</tr></table>
<br>

<p align="center"><input type="button" class="boton" value="IMPRIMIR" onClick="window.open('imprimir_pagina.asp?id=<%=id%>','','scrollbars=yes,resizable=yes,width=400,height=300')">&nbsp;&nbsp;
<%if mid(numeracion_completa,1,1)<>"B" then %>
<input type="button" value="ELIMINAR PÁGINA" class="boton" onClick="location.href='eco_borrarpagina.asp?id=<%=id%>'">
<%else%>
<input type="button" value="ELIMINAR PÁGINA DEFINITIVAMENTE" class="boton" onClick="eco_location.href='eliminarpagina.asp?id=<%=id%>'">
<% end if %>
</p>

</form>
<center>@2012 www.istas.net, All rights reserved</center>
</body>
</html>
