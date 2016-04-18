<!--#include file="web_inicio.asp"-->

<html>
<head>
<title>Listado de códigos CNAE</title>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">

<%

campo = EliminaInyeccionSQL(request("Campo"))
epigrafe = EliminaInyeccionSQL(request("epigrafe"))
subepigrafe = EliminaInyeccionSQL(request("subepigrafe"))
'response.write campo & "///" & epigrafe 
%>

</head>

<body bgcolor="#F4F4F4" topmargin="10" leftmargin="15">

	<SCRIPT LANGUAGE="JScript">
	<!--
	function comprueba_pagina()
	{
			if (document.formulario.epigrafe.value == "0") 
				{ alert('Selecciona un elemento de la lista para acotar la búsqueda'); }
			else	
				{ location.href='listado_cnaes.asp?epigrafe='+document.formulario.epigrafe.value+'&campo=<%=campo%>' ; }
			
	}
	//-->
	</SCRIPT>

<form name="formulario" action="listado_cnaes.asp?" method="POST">
<input type="hidden" name="Campo" value=<%=campo%>>
<p class=negro>Selecciona el código CNAE del listado y pulsa el botón para guardarlo en la ficha</p>
<table class="negroindice">

	<tr><td class="texto" align="left" nowrap>
	Epígrafes principales: </td></tr>
	<tr><td class="texto" align="left"><select class="campo" name="epigrafe" onChange="comprueba_pagina()">
	<% If epigrafe = "" Then %>
		<option value="0" selected>-- Selecciona la actividad de tu empresa </option>
	<% Else %>
		<option value="0">-- Selecciona la actividad de tu empresa </option>
	<% End If
	
		sql = "SELECT * FROM ECO06_LEG_ONLINE_CNAE WHERE len(clasificacion) <= 3 ORDER BY clasificacion"

		Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		set objRecordset2 = OBJConnection.Execute(sql)

		While Not objRecordset2.EOF
			If Len(CStr(objRecordset2("clasificacion"))) = 2 Then
				sangria = "&nbsp;&nbsp;"
			ElseIf Len(CStr(objRecordset2("clasificacion"))) = 3 Then
				sangria = "&nbsp;&nbsp;&nbsp;&nbsp;"
			Else
				sangria = ""
			End If
			strTitulo = objRecordset2("Titulo")
			If Len(CStr(objRecordset2("Titulo"))) > 90 Then
				strTitulo = Left(objRecordset2("Titulo"),87) & " ..."
			End if
			response.write "<option value=" & CStr(objRecordset2("Codigo"))
			If epigrafe = CStr(objRecordset2("Codigo")) Then
				response.write " selected "
			End if
			response.write ">" & sangria & CStr(objRecordset2("Codigo")) & ". " & strTitulo & " </option>" & vbcrlf
			objRecordset2.Movenext
		wend
	%>
	</select>
	</td></tr>
	<tr><td class="texto" align="left" nowrap>
Código CNAE:</td></tr>
	<tr><td class="texto" align="left"><select class="campo" name="subepigrafe"><option value=0 selected>&nbsp;</option>

	<%
		If epigrafe <> "" Then
			sql = "SELECT Clasificacion FROM ECO06_LEG_ONLINE_CNAE WHERE Codigo = '" & epigrafe & "'"
			
			Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			set objRecordset = OBJConnection.Execute(sql)

			If Not objRecordset.EOF Then
				clasificacion = objRecordset("Clasificacion")
				nivel = Len(clasificacion)
				sql = "SELECT * FROM ECO06_LEG_ONLINE_CNAE WHERE Clasificacion LIKE '" & clasificacion & "%' AND len(clasificacion) > " & nivel & " ORDER BY clasificacion"
			Else
				sql = "SELECT * FROM ECO06_LEG_ONLINE_CNAE ORDER BY clasificacion"
				nivel = 0
			End if
		else
			sql = "SELECT * FROM ECO06_LEG_ONLINE_CNAE ORDER BY clasificacion"
			nivel = 0
		End if

		Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		set objRecordset2 = OBJConnection.Execute(sql)

		While Not objRecordset2.EOF
			'If Len(CStr(objRecordset2("clasificacion"))) - nivel = 2 Then
			''	sangria = "&nbsp;&nbsp;"
			'ElseIf Len(CStr(objRecordset2("clasificacion")))  - nivel  = 3 Then
			'	sangria = "&nbsp;&nbsp;&nbsp;&nbsp;"
			'ElseIf Len(CStr(objRecordset2("clasificacion")))  - nivel = 4 Then
			'	sangria = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			'Else
				sangria = ""
			'End If
			strTitulo = objRecordset2("Titulo")
			If Len(CStr(objRecordset2("Titulo"))) > 90 Then
				strTitulo = Left(objRecordset2("Titulo"),87) & " ..."
			End if
			response.write "<option value=" & CStr(objRecordset2("Codigo"))
			'If epigrafe = CStr(objRecordset2("Codigo")) Then
			'	response.write "selected"
			'End if
			response.write ">" & sangria & CStr(objRecordset2("Codigo")) & ". " & strTitulo & " </option>" & vbcrlf
			objRecordset2.Movenext
		wend
	%>
	
	</select>
	</td></tr>
	<tr><td class="texto" align="center" colspan=2>&nbsp;</td></tr>
	<tr><td class="texto" align="center" colspan=2>
	<input type="submit" value="SELECCIONAR CNAE" class="boton" onClick="window.opener.formulario.<%=campo%>_Desc.value=document.formulario.subepigrafe.options[document.formulario.subepigrafe.selectedIndex].text;window.opener.formulario.<%=campo%>.value=document.formulario.subepigrafe.value;window.close();">
	</td></tr>
</table>
</form>
</body>
</html>
