<!--#include file="web_inicio.asp"-->

<%


	texto_busca = limpia(cstr(EliminaInyeccionSQL(request("texto_busca"))))
	Tipo = EliminaInyeccionSQL(request("Tipo"))
	If Tipo = "" Then Tipo = "0"
	subTipo = EliminaInyeccionSQL(request("subTipo"))
	If subTipo = "" Then subTipo = "0"
	ambito = EliminaInyeccionSQL(request("ambito"))
	If ambito = "" Then ambito = "0"
	idautonomia = EliminaInyeccionSQL(request("idautonomia"))
	If idautonomia = "" Then idautonomia = "0"
	If CStr(ambito) = "2" Then subTipo = "0"

	function limpia(texto)
		texto = replace(texto,"'","&#39;")
		limpia = texto
	end Function

	If texto_busca <> "" Then
		textobusqueda = texto_busca
		textobusqueda = ucase(textobusqueda)
		textobusqueda = replace(textobusqueda,"A","[ÁÀAÄ]")
		 textobusqueda = replace(textobusqueda,"E","[ÉÈEË]")
		 textobusqueda = replace(textobusqueda,"I","[ÍIÏÌ]")
		 textobusqueda = replace(textobusqueda,"O","[ÓOÒÖ]")
		 textobusqueda = replace(textobusqueda,"U","[ÚÙUÜ]")
	End If

%>

<html>
<head>
<title>Vigilancia legislativa</title>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
<script>

function brillo(que)
{	que.style.color = "#ff6600";  }

function mate(que)
{	que.style.color = "#000000";  }

function brillo2(que)
{	que.style.color = "#ff6600";  }

function mate2(que)
{	que.style.color = "#AA0000";  }

function ayuda(idpagina)
{	window.status = "Página número "+idpagina; }

function cambianoticia(idnoticia)
{	
	if (idnoticia == 0)
		{
			parent.frames.derecha.location.href='vig_leg_editar.asp?';
		}	
	else
		{
		parent.frames.derecha.location.href='vig_leg_editar.asp?idley='+idnoticia;
		}
	 }

function cambiaejemplar(idnumero)
{
	parent.frames.izquierda.location.href='wi_listado_daphnia.asp?idNumero='+idnumero;
	//document.formulario.submit();
}	

function comprueba_pagina()
{
	enviar = 1;
	
	if (document.formulario.nuevoNumero.value == '' )
	{
		enviar = 0;
		alert('Introduce el número del nuevo Daphnia');	
	}

	if (enviar == 1)
		document.formulario.submit();


}

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
			response.write "addOpt(oCntrl,  " & i & ", '', '0');" & vbcrlf 
			i = i + 1
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

<body bgcolor="#F4F2B8" topmargin="10" leftmargin="15">

<table width="100%" cellpadding="10" cellspacing="0"><tr>
	<td style="border:1px solid #f26522; background: #f6f5e0" align="center">

		<table width="98%" cellpadding="0" cellspacing="2">

		<tr><td class=negro colspan="2" align="center">&nbsp;<b>VIGILANCIA LEGISLATIVA</b></td></tr>
		<tr><td>&nbsp;</td></tr>

		<tr><td colspan="2"  class=negro>

		<form name="formulario" action="vig_leg_listado.asp?" method="POST">

Texto: <BR><input type="text" name="texto_busca" value="<%=texto_busca%>" size="30" class="campo" maxlength="40"><br>

Ámbito: <BR>
<select class="campo" name="ambito" onChange="cambiaAut(document.formulario.idautonomia)">
<option value="0" <% If ambito = "" Or CStr(ambito) = "0" Then Response.write "selected" %>>&nbsp;</option>
<option value="1" <% If CStr(ambito) = "1" Then Response.write "selected" %>>Estatal</option>
<option value="2" <% If CStr(ambito) = "2" Then Response.write "selected" %>>Autonómico</option>
</select><BR>

Comunidad autónoma: <BR>
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

		%></select><BR>

Aspecto ambiental: <BR>
		<select class="campo" name="Tipo" onChange="cambia(document.formulario.Subtipo)">
		<option value="0" <% If Tipo = "" Or CStr(Tipo) = "0" Then Response.write "selected" %>>&nbsp;</option>
		<%		
			
				sql = "SELECT * FROM ECO06_VIG_LEG_TIPOS ORDER BY IdTipo"
			
			Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
			set objRecordset2 = OBJConnection.Execute(sql)

			While Not objRecordset2.EOF
				response.write "<option value=" & CStr(objRecordset2("idtipo"))
				If CStr(Tipo) = CStr(objRecordset2("idtipo")) then
					response.write " selected"
				End if
				response.write ">" & CStr(objRecordset2("nombre_Tipo")) & " </option>" & vbcrlf
				objRecordset2.Movenext
			wend

		%></select><BR>

Aspecto ambiental secundario: <BR>
<select class="campo" name="Subtipo">
		<option value="0" <% If idSubtipo_Ambiental = "" Or CStr(idSubtipo_Ambiental) = "0" Then Response.write "selected" %>>&nbsp;</option>
		<%		
			
				sql = "SELECT * FROM ECO06_VIG_LEG_SUBTIPOS WHERE idTipo = " & CStr(Tipo) & " ORDER BY idTipo, IdsubTipo"
			
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

		%></select><BR>
<BR>

<CENTER><input type="submit" value="BUSCAR" class="boton"> </CENTER>

		</form>
		</td></tr>
		</table>
</td>
</tr>
</table>
<BR>
<table width="98%" cellpadding="5" cellspacing="2">
<%
	'Construimos la cadena de filtrado
	sqlWhere = ""
	If CStr(texto_busca) <> "" Then
		sqlWhere = "WHERE (titulo like '%" & textobusqueda & "%' OR subtitulo like '%" & textobusqueda & "%' OR texto like '%" & textobusqueda & "%') "
	End If
	If CStr(Tipo)	<> "" And CStr(Tipo) <> "0" Then
		If sqlWhere = "" Then
			sqlWhere = "WHERE idTipo_ambiental = " & Tipo 
		Else
			sqlWhere = sqlWhere & " AND idTipo_ambiental = " & Tipo
		End if
	End If
	If CStr(subTipo) <> "" And CStr(subTipo) <> "0" Then
		If sqlWhere = "" Then
			sqlWhere = "WHERE idsubTipo_ambiental = " & subTipo 
		Else
			sqlWhere = sqlWhere & " AND idsubTipo_ambiental = " & subTipo
		End if
	End If	
	If CStr(ambito) <> "" And CStr(ambito) <> "0" Then
		If sqlWhere = "" Then
			sqlWhere = "WHERE ambito = " & ambito 
		Else
			sqlWhere = sqlWhere & " AND ambito = " & ambito
		End if
	End If	
	If CStr(ambito) = "2" then
		If CStr(idautonomia) <> "" And CStr(idautonomia) <> "0" Then
			If sqlWhere = "" Then
				sqlWhere = "WHERE idautonomia = " & idautonomia 
			Else
				sqlWhere = sqlWhere & " AND idautonomia = " & idautonomia
			End if
		End If
	End if

	'response.write sqlWhere & "<BR>"

	If sqlWhere <> "" then

%>

		<tr>
		  <td class="negro" colspan="2" align=center><b>Resultados de la Búsqueda:</b>
		  </td>		
		 </tr>
	<%
		
		j = 1
		i = 0

		sql = "SELECT ECO06_VIG_LEG_LEYES.* FROM ECO06_VIG_LEG_LEYES "
		
		sql = sql & sqlWhere
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		objRecordset.Open sql,OBJConnection,adOpenKeyset
		total_solicitudes = objRecordSet.recordCount

		If total_solicitudes = 0 Then
			strResult = "No se han encontrado leyes con estos parámetros de búsqueda."
		End If
	
	%>
		<tr>
		  <td class="negroindice" colspan="2"><%=strResult%></td>		
		 </tr>
	<%
		registros = 0
		do while not objRecordset.eof
		  registros = registros+1
			
			idley = objRecordset("idley")
			titulo = limpia(cstr(objRecordset("titulo")))
			If Len(titulo) > 90 Then
				titulo = Left(titulo, 90) & " ..."
			End if
		%>	
		<tr>
		  <td class="negroindice" colspan="2">
		  <a style="cursor:hand;" onmouseover=brillo(this) onmouseout=mate(this) onClick="cambianoticia(<%= idley %>)"><% response.write (j & ". " & titulo) %></a></td>		
		 </tr>
	<%
		i = i + 1
		j = j + 1
		objRecordset.movenext
		Loop
	End if	
	%>
	<TR>
		<TD colspan="2">&nbsp;</TD>
	</TR>
	<tr><td colspan="2" align=center>
		<input type="submit" value="NUEVA LEY" class="boton" onClick="cambianoticia(0)">
	</td></tr>

	</table>




</body>
</html>
