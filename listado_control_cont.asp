<!--#include file="web_inicio.asp"-->

<html>
<head>
<title>Listado de sectores en los que hay que prevenir la contaminación</title>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">

<%

epigrafe = EliminaInyeccionSQL(request("epigrafe"))
subepigrafe = EliminaInyeccionSQL(request("subepigrafe"))
'response.write campo & "///" & epigrafe 
%>

<SCRIPT LANGUAGE="JavaScript">

function comprueba_pagina()
{
		if (document.formulario.epigrafe.value == "0") 
			{ alert('Selecciona un elemento de la lista para acotar la búsqueda'); }
		else	
			{ location.href='listado_control_cont.asp?epigrafe='+document.formulario.epigrafe.value; }
		
}

</SCRIPT>

</head>

<body bgcolor="#F4F4F4" topmargin="10" leftmargin="15">

<form name="formulario" action="listado_control_cont.asp?" method="POST">
<p class=negro>Selecciona el tipo de instalación del listado y pulsa el botón Aceptar para guardarlo en la ficha. Si el tipo de instalación de tu empresa no aparece en la lista, selecciona la primera opción: "Ninguna de la lista".</p>
<table class="negroindice">

	<tr><td class="texto" align="left" nowrap>
	Sectores: </td><td class="texto" align="left"><select class="campo" name="epigrafe" onChange="comprueba_pagina();">
	<% If epigrafe = "" Then %>
		<option value="0" selected>-- Selecciona el sector de tu empresa </option>
	<% Else %>
		<option value="0">-- Selecciona el sector de tu empresa </option>
	<% End If
	
		sql = "SELECT * FROM ECO06_LEG_ONLINE_CONTROLCONT WHERE len(clasificacion) = 1 ORDER BY clasificacion"

		Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		set objRecordset2 = OBJConnection.Execute(sql)

		While Not objRecordset2.EOF

			sangria = ""

			strTitulo = objRecordset2("Titulo")
			If Len(CStr(objRecordset2("Titulo"))) > 60 Then
				strTitulo = Left(objRecordset2("Titulo"),57) & " ..."
			End if
			response.write "<option value=" & CStr(objRecordset2("clasificacion"))
			If epigrafe = CStr(objRecordset2("clasificacion")) Then
				response.write " selected "
			End if
			response.write ">" & sangria & strTitulo & " </option>" & vbcrlf
			objRecordset2.Movenext
		wend
	%>
	</select>
	</td></tr>
	<tr><td class="texto" align="left" nowrap>
	Instalaciones:</td>
	<td class="texto" align="left"><select class="campo" name="subepigrafe"><option value=0 selected>&nbsp;</option>
	<option value=1>Ninguna de la lista</option>
	<%
		If epigrafe <> "" Then

			nivel = Len(epigrafe)
			sql = "SELECT * FROM ECO06_LEG_ONLINE_CONTROLCONT WHERE Clasificacion LIKE '" & epigrafe & "%' AND len(clasificacion) > " & nivel & " ORDER BY clasificacion"

		else
			sql = "SELECT * FROM ECO06_LEG_ONLINE_CONTROLCONT WHERE len(Clasificacion) > 1 ORDER BY clasificacion"
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
			'If Len(CStr(objRecordset2("Titulo"))) > 60 Then
			'	strTitulo = Left(objRecordset2("Titulo"),57) & " ..."
			'End if
			response.write "<option value=" & CStr(objRecordset2("clasificacion"))
			'If epigrafe = CStr(objRecordset2("Codigo")) Then
			'	response.write "selected"
			'End if
			response.write ">" & sangria & strTitulo & " </option>" & vbcrlf
			objRecordset2.Movenext
		wend
	%>
	
	</select>
	</td></tr>
	<tr><td class="texto" align="center" colspan=2>&nbsp;</td></tr>
	<tr>
	<td class="texto" align="center" colspan=2>
	<input type="submit" value="ACEPTAR" class="boton" onClick="if (document.formulario.epigrafe.value == 9) {window.opener.document.formulario.Tipo_Inst.value='Ninguna de la lista';window.close();} else {window.opener.document.formulario.Tipo_Inst.value=document.formulario.subepigrafe.options[document.formulario.subepigrafe.selectedIndex].text;window.close();}">
	&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="submit" value="CERRAR" class="boton" onClick="window.opener.document.formulario.Tipo_Inst.value='';window.close();">
	</td>
	</tr>
</table>
</form>
</body>
</html>
