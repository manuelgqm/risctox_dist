<!--#include file="web_inicio.asp"-->

<%
		
	IdLey = EliminaInyeccionSQL(request("IdLey"))
	If IdLey = "" Or CStr(IdLey) = "0" Then
		idTipo =  0
		idSubtipo_Ambiental =  0
		ambito =  0
	End If
	
	'ambito = 2

	'Const adOpenKeyset = 1
	'DIM objConnection	
	'DIM objRecordset
	
	'Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	'Registro con los datos de la página 
	if IdLey <> "" And  CStr(IdLey) <> "0" then 
		sql = "SELECT ECO06_VIG_LEG_LEYES.* FROM ECO06_VIG_LEG_LEYES Where IdLey = " & IdLey
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		objRecordset.Open sql,OBJConnection,adOpenKeyset
		
		If Not objRecordset.EOF then
			If IsNull(objRecordset("IdLey")) Then
				IdLey =  0
			Else
				IdLey =  objRecordset("IdLey")
			End If	
			If IsNull(objRecordset("titulo")) Then
				titulo =  ""
			Else
				titulo =  objRecordset("titulo")
			End if
			If IsNull(objRecordset("subtitulo")) Then
				subtitulo =  ""
			Else
				subtitulo =  objRecordset("subtitulo")
			End if	
			If IsNull(objRecordset("idEnlace")) Then
				idEnlace = 0
			Else
				idEnlace =  objRecordset("idEnlace")
			End If
			If IsNull(objRecordset("texto")) Then
				texto =  ""
			Else
				texto =  objRecordset("texto")
				texto = replace(texto,chr(10),"<br>")
			End If
			If IsNull(objRecordset("idSubtipo_Ambiental")) Then
				idSubtipo_Ambiental =  0
			Else
				idSubtipo_Ambiental =  objRecordset("idSubtipo_Ambiental")
			End If
			If IsNull(objRecordset("idTipo_Ambiental")) Then
				idTipo =  0
			Else
				idTipo =  objRecordset("idTipo_Ambiental")
			End If
			If IsNull(objRecordset("ambito")) Then
				ambito =  0
			Else
				ambito =  objRecordset("ambito")
			End If
			If IsNull(objRecordset("idautonomia")) Then
				idautonomia =  0
			Else
				idautonomia =  objRecordset("idautonomia")
			End If	
			If IsNull(objRecordset("es_LegislacionOnline")) Then
				es_LegislacionOnline =  0
			Else
				es_LegislacionOnline =  objRecordset("es_LegislacionOnline")
			End If
		Else
			idLey = 0
			ambito =  0
			idautonomia =  0
			idTipo =  0
			idSubtipo_Ambiental =  0
			es_LegislacionOnline = 0
		End if
	End if
	
%>	


<html>

<head>
<title>Edición de la ley <%=idley%></title>
<base target="_self">


<script language="javascript" type="text/javascript" src="tinymce/jscripts/tiny_mce/tiny_mce.js"></script>
<script language="javascript" type="text/javascript">
tinyMCE.init({
	mode : "exact",
	elements : "texto,contenido2,contenido3",
	language : "es",
	content_css : "estilos_previsualizacion.css",
	theme : "advanced",
	plugins : "table,paste",
	theme_advanced_toolbar_location : "top",
	theme_advanced_toolbar_align : "left",
	theme_advanced_path_location : "bottom",
	extended_valid_elements : "a[name|href|target|title|onclick],img[class|src|border=0|alt|title|hspace|vspace|width|height|align|onmouseover|onmouseout|name],hr[class|width|size|noshade],font[face|size|color|style],span[class|align|style]",
	external_image_list_url : "example_image_list.js",
	theme_advanced_buttons1 : "bold,italic,underline,strikethrough,separator,justifyleft,justifycenter,justifyright,justifyfull,separator,cut,copy,pastetext,separator,tablecontrols",
	theme_advanced_buttons2 : "styleselect,separator,undo,redo,separator,bullist,numlist,outdent,indent,separator,sub,sup,separator,charmap,cleanup,hr,removeformat,help,code",
	theme_advanced_buttons3 : ""


});
</script>

<%
response.write "<script language='JavaScript'>" & vbcrlf

response.write "function addOpt(oCntrl, iPos, sTxt, sVal){ " & vbcrlf
response.write "var selOpcion=new Option(sTxt, sVal); " & vbcrlf
response.write "eval(oCntrl.options[iPos]=selOpcion); " & vbcrlf
response.write "} " & vbcrlf

response.write "function cambiaAut(oCntrl){ " & vbcrlf
response.write "while (oCntrl.length) oCntrl.remove(0); " & vbcrlf
response.write "switch (document.formulario.ambito.selectedIndex){ " & vbcrlf

		response.write "case 0:" & vbcrlf
		response.write "addOpt(oCntrl,  0, '', '0');" & vbcrlf 
		response.write "break; " & vbcrlf
		response.write "case 1:" & vbcrlf
		response.write "addOpt(oCntrl,  0, '', '0');" & vbcrlf 
		response.write "break; " & vbcrlf
		response.write "case 2:" & vbcrlf

		sql = "SELECT * FROM ECOINFORMAS_VALORES WHERE Grupo = '032' ORDER BY cast(subgrupo as int)"
		
		Set objRecordset4 = Server.CreateObject ("ADODB.Recordset")
		set objRecordset4 = OBJConnection.Execute(sql)
		
		i = 0
		While Not objRecordset4.EOF
			response.write "addOpt(oCntrl,  " & i & ", '" & CStr(objRecordset4("desc1")) & "', '" & CStr(objRecordset4("subgrupo")) & "');" & vbcrlf 
			i = i + 1
			objRecordset4.Movenext
		wend

		response.write "break; " & vbcrlf

	 response.write "}  " & vbcrlf
	response.write "}  " & vbcrlf
  'response.write "</script> " & vbcrlf

response.write "function cambia(oCntrl){ " & vbcrlf
response.write "while (oCntrl.length) oCntrl.remove(0); " & vbcrlf
response.write "if (document.formulario.ambito.selectedIndex == 2){ " & vbcrlf

		response.write "addOpt(oCntrl,  0, '', '0');" & vbcrlf 

response.write "} else { " & vbcrlf
response.write "switch (document.formulario.Tipo.selectedIndex){ " & vbcrlf

	sql = "SELECT * FROM ECO06_VIG_LEG_SUBTIPOS ORDER BY idTipo, IdsubTipo"
			
	Set objRecordset3 = Server.CreateObject ("ADODB.Recordset")
	set objRecordset3 = OBJConnection.Execute(sql)
	intTipo = 0
	i = -1
	While Not objRecordset3.EOF
		If CStr(intTipo) <> CStr(objRecordset3("idtipo")) Then
			If i > -1 Then
				response.write "break; " & vbcrlf
			End if
			response.write "case " & CStr(objRecordset3("idtipo")) & ":" & vbcrlf
			i = 0
		End If
		response.write "addOpt(oCntrl,  " & i & ", '" & CStr(objRecordset3("nombre_subtipo")) & "', '" & CStr(objRecordset3("idsubtipo")) & "');" & vbcrlf 
		intTipo = objRecordset3("idtipo")
		i = i + 1 
		objRecordset3.Movenext
	Wend
	If i > -1 Then
		response.write "break; " & vbcrlf
	End if
		response.write "}  " & vbcrlf
	 response.write "}  " & vbcrlf
	response.write "}  " & vbcrlf
  response.write "</script> " & vbcrlf
%>

</head>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
<body bgcolor="#F4F2B8" topmargin="20" leftmargin="20">
<script src="valida_fecha.js"></script>
<script LANGUAGE="JScript">
<!--
// 
    ua=navigator.userAgent; v=navigator.appVersion.substring(0,1);
    if ((ua.lastIndexOf("MSIE")!=-1) && (v!='1') && (v!='2') && (v!='3')) { document.body.onmouseover=Brillo; document.body.onmouseout=Mate }
    function Brillo() { src = event.toElement; if (src.tagName == "A") { src.antes = src.style.color; src.style.color ="#F09450" } }
    function Mate() { src=event.fromElement; if (src.tagName == "A") { src.style.color = src.antes } }



function comprueba_pagina()
{
	enviar = 1;
	
	if (document.formulario.titulo.value == "")
	{
		enviar = 0;
		alert('Introduce el título de la ley');	
	}

	if (document.formulario.Tipo.value == "0" && enviar == 1 )
	{
		enviar = 0;
		alert('Introduce el aspecto ambiental de la ley');	
	}

	if (document.formulario.Subtipo.value == "0" && document.formulario.ambito.value == "1" && enviar == 1)
	{
		enviar = 0;
		alert('Introduce el aspecto ambiental secundario de la ley');	
	}

	if (document.formulario.ambito.value == "0" && enviar == 1)
	{
		enviar = 0;
		alert('Introduce el ámbito de la ley');	
	}

		if (document.formulario.ambito.value == "2" && document.formulario.idautonomia.value == "0" && enviar == 1)
	{
		enviar = 0;
		alert('El ámbito de la ley es autonómico, pero no hay ninguna comunidad autónoma seleccionada');	
	}

	if (enviar == 1)
		document.formulario.submit();


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



//-->
</script>

<!--------------------------------------------------------------------CONTENIDO--> 
<form name="formulario" ACTION="vig_leg_modificar.asp?idley=<%=idLey%>" METHOD="POST">

<!-- CASTELLANO -->
<table width="100%" cellpadding="0" cellspacing="0"><tr><td class="celda" align="center">
<BR>
<table width="95%"><tr><td class="cue_celda"><b>Título ley:</b></td></tr></table>
<input type="text" name="titulo" value="<%=titulo%>" size="88" class="campo" maxlength="255"><br>
<table width="95%"><tr><td class="cue_celda"><b>Subtítulo ley:</b></td></tr></table>
<input type="text" name="subtitulo" value="<%=subtitulo%>" size="88" class="campo" maxlength="255"><br>
<table width="95%"><tr><td class="cue_celda"><b>Texto:</b></td></tr></table>
<textarea name="texto" rows="16" cols="88" class="campo"><%=texto%></textarea><br>
<!--------------------------------------------------------------------ENLACE --> 
<table width="95%"><tr><td class="cue_celda"><b>Insertar enlace BV:</b></td></tr></table>
<table width="95%"><tr><td class="cue_fuente" align=center>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="text" name="enlaces" value="<%=idEnlace%>" size="10" class="campo" maxlength="10">&nbsp;
	<input type="button" value="VER BV" class="boton" onClick="window.open('listado_enlaces.asp?campo=enlaces','listado_enlaces','width=300,height=300,resizable=yes,scrollbars=yes,menubar=yes')">
</td>
</tr>
</table>
<!--------------------------------------------------------------------TIPO LEY --> 
<table width="95%"><tr><td class="cue_celda"><b>Información adicional de la ley:</b></td></tr></table>

<table width="95%">
<!-- <td class="cue_fuente">Tipo:&nbsp;
<select class="campo" name="tipo">
<option value="1" class="azulindice" <%if tipo=1 then response.write "selected"%>>Columna</option>
<option value="2" class="azulindice" <%if tipo=2 then response.write "selected"%>>Dos columnas</option>
<option value="3" class="azulindice" <%if tipo=3 then response.write "selected"%>>Carrusel</option>
</select>&nbsp;
</td> -->
<tr>
<td align="right" class="cue_fuente" width=40%>Ámbito: &nbsp;</td>
<td align="left" class="cue_fuente">
<select class="campo" name="ambito" onChange="cambiaAut(document.formulario.idautonomia)">
<option value="0" <% If ambito = "" Or CStr(ambito) = "0" Then Response.write "selected" %>>&nbsp;</option>
<option value="1" <% If CStr(ambito) = "1" Then Response.write "selected" %>>Estatal</option>
<option value="2" <% If CStr(ambito) = "2" Then Response.write "selected" %>>Autonómico</option>
</select>
</td>
</tr>
<tr>
<% If ambito = 2 Then %>
<td align="right" class="cue_fuente" width=40%>Comunidad autónoma: &nbsp;</td>
<td align="left" class="cue_fuente">
		<select class="campo" name="idautonomia">
		<option value="0" <% If idautonomia = "" Or CStr(idautonomia) = "0" Then Response.write "selected" %>>&nbsp;</option>
		<%		
			
				sql = "SELECT * FROM ECOINFORMAS_VALORES WHERE Grupo = '032' ORDER BY cast(subgrupo as int)"
			
			Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
			set objRecordset2 = OBJConnection.Execute(sql)

			While Not objRecordset2.EOF
				response.write "<option value=" & CStr(objRecordset2("subgrupo"))
				If CStr(idautonomia) = CStr(objRecordset2("subgrupo")) then
					response.write " selected"
				End if
				response.write ">" & CStr(objRecordset2("desc1")) & " </option>" & vbcrlf
				objRecordset2.Movenext
			wend

		%></select>
</td>
<% Else %>
<td align="right" class="cue_fuente" width=40%>Comunidad autónoma: &nbsp;</td>
<td align="left" class="cue_fuente">
		<select class="campo" name="idautonomia">
		<option value="0" <% Response.write "selected" %>>&nbsp;</option>
		</select>
</td>
<% End If %>

</tr>
<tr>
<td align="right" class="cue_fuente" width=40%>Aspecto ambiental: &nbsp;</td><td align="left" class="cue_fuente">
		<select class="campo" name="Tipo" onChange="cambia(document.formulario.Subtipo)">
		<option value="0" <% If idTipo = "" Or CStr(idTipo) = "0" Then Response.write "selected" %>>&nbsp;</option>
		<%		
			
				sql = "SELECT * FROM ECO06_VIG_LEG_TIPOS ORDER BY IdTipo"
			
			Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
			set objRecordset2 = OBJConnection.Execute(sql)

			While Not objRecordset2.EOF
				response.write "<option value=" & CStr(objRecordset2("idtipo"))
				If CStr(idTipo) = CStr(objRecordset2("idtipo")) then
					response.write " selected"
				End if
				response.write ">" & CStr(objRecordset2("nombre_Tipo")) & " </option>" & vbcrlf
				objRecordset2.Movenext
			wend

		%></select>
</td>
</tr>
<tr>
<td align="right" class="cue_fuente" width=40%>Aspecto ambiental secundario: &nbsp;</td><td align="left" class="cue_fuente">
		<select class="campo" name="Subtipo">
		<option value="0" <% If idSubtipo_Ambiental = "" Or CStr(idSubtipo_Ambiental) = "0" Then Response.write "selected" %>>&nbsp;</option>
		<%		
			
				sql = "SELECT * FROM ECO06_VIG_LEG_SUBTIPOS WHERE idTipo = " & CStr(idTipo) & " ORDER BY idTipo, IdsubTipo"
			
			Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
			set objRecordset2 = OBJConnection.Execute(sql)

			While Not objRecordset2.EOF
				response.write "<option value=" & CStr(objRecordset2("idsubtipo"))
				If CStr(idSubtipo_Ambiental) = CStr(objRecordset2("idsubtipo")) then
					response.write " selected"
				End if
				response.write ">" & CStr(objRecordset2("nombre_subtipo")) & " </option>" & vbcrlf
				objRecordset2.Movenext
			wend

		%></select>
</td>
</tr>
<tr>
<td align="right" class="cue_fuente" width=40%>Pertenece a legislación online: &nbsp;</td><td align="left" class="cue_fuente">
		<select class="campo" name="es_LegislacionOnline">
		<option value="0" <% If es_LegislacionOnline = "" Or CStr(es_LegislacionOnline) = "0" Then Response.write "selected" %>>No</option>
		<option value="1" <% If CStr(es_LegislacionOnline) = "1" Then Response.write "selected" %>>Sí</option>		
</select>
</td>
</tr>
</table>

<BR>

<% If idLey <> "" And idLey <> "0" Then %>
<table align="center"><tr><td><input type="button" value="GUARDAR CAMBIOS" class="boton" onClick="comprueba_pagina(1)">&nbsp;&nbsp;</td><td>&nbsp;&nbsp;<input type="button" value="BORRAR" class="boton"
onclick="if (confirm('Se borrará la ley. ¿Deseas continuar?')) {location.href='vig_leg_borrar.asp?idley=<%=idley%>';}"></td></tr></table>
<br>
<% Else %>
<table align="center"><tr><td><input type="button" value="GUARDAR CAMBIOS" class="boton" onClick="comprueba_pagina(1)"></td></tr></table>
<br>
<% End If %>
<!-- <p align="center"><input type="button" class="boton" value="VISTA PREVIA" onclick="window.open('daphnia.asp?articulo=<%=idarticulo%>','vistaprevia','scrollbars=yes,resizable=yes,status=yes,toolbar=yes,location=yes,width=640,height=480')"></p>


<br> -->
<br>
</td>
</tr>

</table>
</form>

<% 'end if%>
</body>
</html>
