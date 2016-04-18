<!--#include file="eco_conexion.asp"-->
<%
	'Función: Muestra un formulario para editar una página.
	
	'Parámetros de entrada:
		'id: Identificador de la página
		
	id = clng(EliminaInyeccionSQL(request("id")))

	'Registro con los datos de la página 
	sql = "SELECT numeracion FROM WEBISTAS_PAGINAS WHERE idpagina="&id
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	numeracion = objRecordSet("numeracion")
	objRecordset.close
	
	'sql = "SELECT idpagina FROM WEBISTAS_PAGINAS WHERE numeracion LIKE '"&numeracion&"%' AND len(numeracion)="&cstr(len(numeracion)+1)&" ORDER BY numeracion"
	'Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	'objRecordset.Open sql,OBJConnection,adOpenKeyset
	'hijos = objrecordset.recordcount
	
	sql = "SELECT numeracion FROM WEBISTAS_PAGINAS WHERE numeracion LIKE '"&numeracion&"%' AND len(numeracion)="&cstr(len(numeracion)+1)&" ORDER BY numeracion DESC"
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	objRecordset.Open sql,OBJConnection,adOpenKeyset
	if objrecordset.eof then
		hijos = 0
	else
		letra_hijomayor = objrecordset("numeracion")
		hijos = asc(mid(letra_hijomayor,(len(numeracion)+1),1))-64
	end if


%>	


<html>

<head>
<title>Nueva subpágina de la página <%=id%></title>
<base target="_self">

</head>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
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
	
	if ((document.formulario.tipo.value == "1")||(document.formulario.tipo.value == "2"))
		if (document.formulario.titulo.value != "")
			{ document.formulario.submit(); }
		else
			{ alert('Introduce el título'); }
	else
		if (document.formulario.destino.value == "")
			{ alert('Introduce el destino'); }
		else
			{ document.formulario.submit(); }
}

function formato(codigo)
{
	if (document.selection.createRange().text!='')
	{ document.selection.createRange().text = '<'+codigo+'>'+document.selection.createRange().text+'</'+codigo+'>' }
	else
	{ alert('Hay que seleccionar un texto para poder aplicarle formato') }
}

function formatoenlace(codigo)
{
	if ((document.formulario.enlace_url.value!='') && (document.selection.createRange().text!=''))
	{ document.selection.createRange().text = '<'+codigo+'='+document.formulario.enlace_url.value+'>'+document.selection.createRange().text+'</'+codigo+'>' }
else
	{ alert('Hay que seleccionar un texto y escribir su enlace antes de presionar este botón') }
}


//-->
</script>
<form name="formulario" ACTION="eco_nuevasubpagina.asp" METHOD="POST">
<!--------------------------------------------------------------------CLASIFICACIÓN--> 
<% for i=2 to 10
	valor = ""
	if len(numeracion)>=i then 
		valor = cstr(asc(mid(numeracion,i,1))-64)
		numeracioncompleta = numeracioncompleta&valor&"."
	end if
   next %>
<table><tr><td class="cue_celda"><b>Nueva página <%=numeracioncompleta&cstr(hijos+1)%>:</b></td></tr></table>
<br><br>

<!--------------------------------------------------------------------TÍTULO--> 
<table><tr><td class="cue_celda"><b>Título:</b></td></tr></table>
<input type="hidden" name="numeracion" value="<%=numeracion&chr(hijos+65)%>">
<input type="text" name="titulo" size="100" class="campo" maxlength="255">
<br><br>

<!--------------------------------------------------------------------TIPO--> 
<table><tr><td class="cue_celda"><b>Tipo:</b></td>
<td class="cue_fuente"><select class="campo" name="tipo">
<option value="1" class="azulindice" >Página HTML (no aplica los códigos)</option>
<option value="2" class="azulindice" selected>Página normal</option>
<option value="3" class="azulindice" >Otra página ya existente</option>
<option value="4" class="azulindice" >Ficha de recursos para pymes</option>
<option value="5" class="azulindice" >Carpeta de fichas de rpp</option>
<option value="6" class="azulindice" >Carpeta de registros de la bv</option>
</select>&nbsp;
Destino:&nbsp;<input type="text" name="destino" value="<%=destino%>" size="4" class="campo" maxlength="5"></td>
<td class="cue_fuente">
&nbsp;&nbsp;&nbsp;Visible&nbsp;<select class="campo" name="visible">
<option value="1">S?/option>
<option value="2">No</option>
</select></td>
</tr></table>
<br><br>

<!--------------------------------------------------------------------CONTENIDO--> 
<table class="cue_celda" width="85%"><tr><td><b>Contenido:</b></td>
<td align="right">
<img src="negrita.gif" onClick="formato('b')" alt="negrita" style="cursor:hand">&nbsp;
<img src="cursiva.gif" onClick="formato('i')" alt="cursiva" style="cursor:hand">&nbsp;
<img src="subrayado.gif" onClick="formato('u')" alt="subrayada" style="cursor:hand">&nbsp;
<img src="titulo.gif" onClick="formato('t')" alt="título" style="cursor:hand">&nbsp;
<img src="subtit.gif" onClick="formato('st')" alt="subtítulo" style="cursor:hand">&nbsp;
<img src="parrafo.gif" onClick="formato('pd')" alt="parrafo destacado" style="cursor:hand">&nbsp;
</td>
</tr></table>
<textarea rows="25" name="pagina" cols="100" class="campo"><%=pagina%>
</textarea><br>


<p align="center">
<input type="button" value="GUARDAR NUEVA PÁGINA" class="boton" onClick="comprueba_pagina()">
<input type="button" value="VOLVER SIN GUARDAR" class="boton" onClick="location.href='editarpagina.asp?id=<%=id%>'">&nbsp;&nbsp;
</p>

</form>
<center><a href="http://www.toporologi.org" title="replica,rolex,orologi">rolex</a>@2012 www.istas.net, All rights reserved</center>

</body>
</html>
