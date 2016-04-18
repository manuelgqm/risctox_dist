<!--#include file="web_inicio.asp"-->
<%
 	'Const adOpenKeyset = 1
	'DIM objConnection	
	'DIM objRecordset
	
	'Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
		
	'----- Si es restringida y no estás identificado no puedes entrar
	'if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	
	CNAE_1 = request("CNAE_1")
	CNAE_2 = request("CNAE_2")
	CNAE_3 = request("CNAE_3")
	CNAE_1_Desc = request("CNAE_1_Desc")
	CNAE_2_Desc = request("CNAE_2_Desc")
	CNAE_3_Desc = request("CNAE_3_Desc")
	idautonomia = request("idautonomia")
	empresas = request("empresas")
	Nombre_Empresa = request("Nombre_Empresa")
	Direccion = request("Direccion")
	Poblacion = request("Poblacion")
	CP = request("CP")
	Grupo_1 = request("Grupo_1")
	Grupo_2 = request("Grupo_2")
	Grupo_3 = request("Grupo_3")
	Grupo_4 = request("Grupo_4")
	Grupo_5 = request("Grupo_5")
	Grupo_6 = request("Grupo_6")
	Tipo_Inst = request("Tipo_Inst")

	If empresas <> "" And empresas <> "0" Then
		es_GEI = 1
	Else
		es_GEI = 0
	End If
			
	function limpia(texto)
		texto = replace(texto,"'","&#39;")
		limpia = texto
	end Function

	Function limpiaTextoLey(texto)
		texto = replace(texto,"'","&#39;")
		texto = Replace(texto,"<P>"," ")
		texto = Replace(texto,"</P>"," ")
		texto = Replace(texto,"<p>"," ")
		texto = Replace(texto,"</p>"," ")
		texto = Trim(texto)
		limpiaTextoLey = texto
	end Function

	Function EscribeLey(strTitulo,strSubTitulo,strTexto,strEnlace)
		'Esta funcion es para poder cambiar facilmente cómo se muestran las leyes
		'Devuelve Ok si todo va bien 
		If strTitulo <> "" Then
			response.write "<b>- " & strTitulo & "</b><BR>"
		End If
		If strSubTitulo <> "" Then
			response.write "&nbsp;&nbsp;" & strSubTitulo & "<BR>"
		End If
		If strTexto <> "" Then
			response.write "&nbsp;&nbsp;" & strTexto & "<BR>"
		End If
		If strEnlace <> "" Then
			response.write "&nbsp;&nbsp;<A HREF='abreenlace.asp?idenlace=" & strEnlace & "' target='_blank'>Enlace a la ley: http://www.istas.net/ecoinformas/web/abreenlace.asp?idenlace=" & strEnlace & "</A><BR><BR>"
		End If
		EscribeLey = "ok"
	End Function
	
	Function FraseEmision(GRUPO_CNAE)
		frase_1 = "<b>Emisiones atmosféricas:</b><BR>&nbsp;Autorización departamento agricultura. <a href='#notas'>Ver nota (3)</a>.<br><BR>"
		frase_2 = "<b>Emisiones atmosféricas:</b><BR>&nbsp;Autorización departamento obras públicas. <a href='#notas'>Ver nota (3)</a>.<br><BR>"
		strResultado = "<b>Emisiones atmosféricas:</b><BR>&nbsp;Autorización departamento industria. <a href='#notas'>Ver nota (3)</a>.<br><BR>"
		
		Select Case GRUPO_CNAE
			Case "AA","01","02","05"
				strResultado = frase_1
			Case "CA","10","11","12", "CB", "13","14"
				strResultado = frase_2
		End Select
		
		FraseEmision = strResultado
	End Function 

	If texto_busca <> "" Then
		textobusqueda = texto_busca
		textobusqueda = ucase(textobusqueda)
		textobusqueda = replace(textobusqueda,"A","[ÁÀAÄ]")
		 textobusqueda = replace(textobusqueda,"E","[ÉÈEË]")
		 textobusqueda = replace(textobusqueda,"I","[ÍIÏÌ]")
		 textobusqueda = replace(textobusqueda,"O","[ÓOÒÖ]")
		 textobusqueda = replace(textobusqueda,"U","[ÚÙUÜ]")
	End If

	'Rellenamos ahora estos textos para no hacerlo luego varias veces
	sql = "SELECT * FROM ECO06_LEG_ONLINE_LEYES_BASICAS"
				
		Set objRecordset5 = Server.CreateObject ("ADODB.Recordset")
		set objRecordset5 = OBJConnection.Execute(sql)	

		While Not objRecordset5.EOF
			Select Case objRecordset5("Aspecto")
				Case "GEI"
					strGEI = objRecordset5("Texto")
				Case "COVS"
					strCOVS = objRecordset5("Texto")
				Case "COP"
					strCOP = objRecordset5("Texto")
				Case "SEVESO"
					strSEVESO = objRecordset5("Texto")
				Case "ContAtmosferica"
					strContAtmosferica = objRecordset5("Texto")
			End Select 
			objRecordset5.Movenext
		Wend 
	
	FUNCTION vistaprevia(texto)
		texto = replace(texto,chr(13),"<br>")
		texto = replace(texto,"'","&#39;")
		texto = replace(texto,"<v1>","<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v2>","&nbsp;&nbsp;<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v3>","&nbsp;&nbsp;&nbsp;&nbsp;<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v4>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v5>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<v6>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=imagenes/vineta.gif>&nbsp;")
		texto = replace(texto,"<pag=","<a href=index.asp?idpagina=")
		texto = replace(texto,"</pag>","</a>")
		texto = replace(texto,"<e=","<a target=_blank href=abreenlace.asp?idenlace=")
		texto = replace(texto,"<er=","<a target=_blank href=abreenlacer.asp?idenlace=")
		texto = replace(texto,"</e>","</a>")
		texto = replace(texto,"<t>","<font class=titulo"&(seccion)&">")
		texto = replace(texto,"</t>","</font>")
		texto = replace(texto,"<st>","<font class=subtitulo"&(seccion)&">")
		texto = replace(texto,"</st>","</font>")
		texto = replace(texto,"<pd>","<table width=95% align=center cellpadding=10 cellspacing=0 class=tabla><tr><td>")
		texto = replace(texto,"</pd>","</td><td valign=top align=center><img src=pd.gif></td></tr></table>")
		vistaprevia = texto
		
	END FUNCTION

	function contenga(codigo)
		texto_sql = "("
		for n=1 to len(codigo)-1
			texto_sql = texto_sql & " ascii(substring(numeracion,"&cstr(n)&",1))=" & asc(mid(codigo,n,1)) & " AND "
		next
		texto_sql = texto_sql & " ascii(substring(numeracion,"&cstr(len(codigo))&",1))=" & asc(mid(codigo,len(codigo),1)) & ")"
		contenga = texto_sql
		'response.write texto_sql & "<br>"
	end function

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Imprime: Legislación Online</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multimèdia" />
<meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<link rel="stylesheet" type="text/css" href="estructura.css"  />


</head>
<body onload="print()">
<div id="caja_imprimir">
	<div class="texto">
	<%  'Si el usuario ha consultado mostramos los resultados
		If (empresas <> "" And empresas <> "0") Or Nombre_Empresa <> "" Then %>
		<BR>
			<P style="padding-left:5px;" class="titulo_imprimir">Resultados de la consulta de legislación onLine:</P>
			<!-- <B><U>Autorizaciones:</U></B><BR><BR> -->
			<IMG SRC="imagenes/leg_online_autorizacion.gif" WIDTH="146" HEIGHT="32" BORDER="0" ALT=""><BR><BR>
		<%
			response.write "<TABLE cellpadding='5' style='border:1px solid #000000;'><tr><td>"
			'Excepcion, si se elige listado de instalaciones un tipo de empresa
			'se elige idopcion 20, si no 30 
			If CStr(Tipo_Inst) <> "" And CStr(Tipo_Inst) <> "Ninguna de la lista" Then
				intControlCont = "20"
			Else
				intControlCont = "30"
			End If
	
			'Excepcion: el usuario dice que no hay calderas, pero su CNAE indica que sus 
			'actividades son peligrosas para contaminación atmosferica. 
			'Actuamos como si hubiera dicho que sí ...
			If CStr(Grupo_3) = "120" Then
				sqlWhere = " Codigo = '" & CNAE_1 & "'"
				If CNAE_2 <> "" Then
					sqlWhere = sqlWhere & " OR Codigo = '" & CNAE_2 & "'"
				End If
				If CNAE_3 <> "" Then 
					sqlWhere = sqlWhere & " OR Codigo = '" & CNAE_3 & "'"
				End If  
				sql = "SELECT Afecta_ContAtmosferica FROM ECO06_LEG_ONLINE_CNAE WHERE " & sqlWhere & " ORDER BY Afecta_ContAtmosferica DESC"

				Set objRecordset5 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset5 = OBJConnection.Execute(sql)	

				If Not objRecordset5.EOF Then
					If CStr(objRecordset5("Afecta_ContAtmosferica")) <> "0" Then
						Grupo_3 = "110"
					End if
				End if
			End If
			
			'Primero ponemos las autorizaciones, segun las opciones que selecciona el usuario
			'La opcion 10 es la de urbanismo, sale simepre
			sqlWhere = "(IdOpcion = 10 or IdOpcion = '" & Grupo_1 & "' or IdOpcion = '" & Grupo_2 & "' or IdOpcion = '" & Grupo_3 & "' or IdOpcion = '" & Grupo_4 & "' or IdOpcion = '" & Grupo_5 & "' or IdOpcion = '" & Grupo_6 & "' or IdOpcion = '" & intControlCont & "')"

			sql = "SELECT * FROM ECO06_LEG_ONLINE_AUTORIZACIONES WHERE " & sqlWhere & " ORDER BY IdOpcion"

			Set objRecordset4 = Server.CreateObject ("ADODB.Recordset")
			set objRecordset4 = OBJConnection.Execute(sql)			

			While Not objRecordset4.EOF
				'Hay una excepcion para emisiones a la atmosfera
				If objRecordset4("IdOpcion") <> 110 And objRecordset4("IdOpcion") <> 130 and objRecordset4("IdOpcion") <> 140 Then
					response.write "<b>" & objRecordset4("Aspecto_Ambiental") & ":</b><BR>&nbsp;" & objRecordset4("Texto_Autorizacion") & "<br><BR>"
				Else
					'If Len(CNAE_1) = 1 Then
					'	response.write FraseEmision(CNAE_1 & "A")
					'else
					'	response.write FraseEmision(Left(CNAE_1,2))
					'End If

					'Hacemos aqui las excepciones a las autorizaciones (COVS y emisiones atmosfericas)
					If objRecordset4("IdOpcion") = 110 then
						response.write "<b>Emisiones atmosféricas:</b><BR>&nbsp;" 
						response.write objRecordset4("Texto_Autorizacion")
						response.write "<BR><BR>"
					ElseIf objRecordset4("IdOpcion") = 130 Then
						If intControlCont = "20" then
							response.write "<b>Emisiones COV:</b><BR>&nbsp;" 
							response.write "La Autorización ambiental integrada deberá incluir los Valores Límite de Emisión o los sistemas de reducción de emisiones de COVs. <a href='#notas'>Ver nota (4)</a>."
							response.write "<BR><BR>"						
						Else
							response.write "<b>Emisiones COV:</b><BR>&nbsp;" 
							response.write "Notificación para su inscripción en el Registro de instalaciones emisoras de COVs. <a href='#notas'>Ver nota (5)</a>."
							response.write "<BR><BR>"							
						End If 
					ElseIf objRecordset4("IdOpcion") = 140 then
						If intControlCont = "20" then
							response.write "<b>Emisiones COV:</b></td></tr>" 
							response.write "La Autorización ambiental integrada deberá incluir los Valores Límite de Emisión o los sistemas de reducción de emisiones de COVs. <a href='#notas'>Ver nota (4)</a>."
							response.write "<BR><BR>"					
						Else
							response.write "<b>Emisiones COV:</b><BR>&nbsp;" 
							response.write "Notificación para su inscripción en el Registro de instalaciones emisoras de COVs. <a href='#notas'>Ver nota (5)</a>."
							response.write "<BR><BR>"							
						End If 
					End If 
				End If 
				objRecordset4.Movenext
			wend

			'Si está afectado por la GEI, se anota ahora
			If CStr(empresas) <> "" And CStr(empresas) <> "0" then

				response.write "<b>Emisión de gases de efecto invernadero:</b><BR>&nbsp;" & strGEI & "<br>"

			''''' HAsta que no esten las tn asignadas a cada empresa nop mostramos nada '''''''''''''''''''''''''''''''''''''''''''''''
			'	
			'	sql = "SELECT * FROM ECO06_LEG_ONLINE_GEI WHERE IdInstalacion = " & Empresas
			'	
			'	Set objRecordset4 = Server.CreateObject ("ADODB.Recordset")
			'	set objRecordset4 = OBJConnection.Execute(sql)	
			'
			'	If Not objRecordset4.EOF Then
			'		response.write "<b>Cantidad asignada (Tn de CO2 por año):</b>&nbsp;" & objRecordset4("Anyo_1") & ": " & 'objRecordset4("Cantidad_1") & ";&nbsp; " & objRecordset4("Anyo_2") & ": " & objRecordset4("Cantidad_2") & ";&nbsp; " & 'objRecordset4("Anyo_3") & ": " & objRecordset4("Cantidad_3") & ".<BR>"
			'	End If 
			'
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			End If
			response.write "</td></tr></table>"
		%>
			<BR>
			<!-- <B><U>Legislación básica:</U></B><BR><BR> -->
			<BR><IMG SRC="imagenes/leg_online_legislacion.gif" WIDTH="197" HEIGHT="32" BORDER="0" ALT=""><BR><BR>
		<%
			response.write "<TABLE cellpadding='5' style='border:1px solid #000000;'><tr><td>"

			actividad_afectada = 0

			sql = "SELECT * FROM ECO06_LEG_ONLINE_CNAE"		
			sqlWhere = " WHERE Codigo = '" & CNAE_1 & "'"
			If CNAE_2 <> "" Then
				sqlWhere = sqlWhere & " OR Codigo = '" & CNAE_2 & "'"
			End If
			If CNAE_3 <> "" Then 
				sqlWhere = sqlWhere & " OR Codigo = '" & CNAE_3 & "'"
			End If 
			sql = sql & sqlWhere 
			
			Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
			set objRecordset2 = OBJConnection.Execute(sql)	
		
			Dim miTabla
			If Not objRecordset2.EOF then
				miTabla = objRecordset2.GetRows 
			End If 			

			'Primero Prevención y control de la contaminación
			If CStr(Tipo_Inst) <> "" And CStr(Tipo_Inst) <> "Ninguna de la lista"  Then
				'Buscamos en Vig. legislativa aspecto ambiental : AUTORIZACIONES AMBIENTALES 
				'Subaspecto ambiental: AUTORIZACIÓN AMBIENTAL INTEGRADA (tipo 7, subtipo 28)
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE IdTipo_Ambiental = 7 AND (IdSubtipo_Ambiental = 28 OR IdAutonomia = " & idautonomia & ") AND Es_LegislacionOnline = 1 ORDER BY ambito"

			   		   set objRecordset = Server.CreateObject ("ADODB.Recordset")
			   		   set objRecordset = OBJConnection.Execute(sql)

				response.write "<b>Prevención y control integrados de la contaminación:</b><BR>"

				if Not objRecordset.EOF Then
				
					While Not objRecordset.EOF	   
						If IsNull(objRecordset("titulo")) Then
							strTituloLey = ""
						Else
							strTituloLey = objRecordset("titulo")
						End if
						If IsNull(objRecordset("subtitulo")) Then
							strSubTituloLey = ""
						Else
							strSubTituloLey = objRecordset("subtitulo")
						End If
						If IsNull(objRecordset("texto")) Then
							strTextoLey = ""
						Else
							strTextoLey = limpiaTextoLey(objRecordset("texto"))
						End If
						If IsNull(objRecordset("idenlace")) Then
							strEnlaceLey = ""
						Else
							strEnlaceLey = CStr(objRecordset("idenlace"))
						End If			
						strOK = EscribeLey(strTituloLey,strSubTituloLey,strTextoLey,strEnlaceLey)
					objRecordset.MoveNext
					Wend	
				Else
					response.write "&nbsp;&nbsp;La actividad de la empresa no está afectada por la legislación ambiental en este aspecto.<BR>"		
				End If
			Else
				response.write "<b>Prevención y control integrados de la contaminación:</b><BR>"
				response.write "&nbsp;&nbsp;La actividad de la empresa no está afectada por la legislación ambiental en este aspecto.<BR>"
			End if

			'Buscamos en Vig. legislativa aspecto ambiental Residuos:
			If CStr(grupo_1) = "50" Or CStr(grupo_1) = "60" then
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE ((IdTipo_Ambiental = 1 AND IdAutonomia = " & idautonomia & ") OR (IdTipo_Ambiental = 1 AND (IdSubtipo_Ambiental = 13 OR IdSubtipo_Ambiental = 14))) AND Es_LegislacionOnline = 1 ORDER BY ambito"
			Else
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE idenlace = 2946 AND Es_LegislacionOnline = 1"
			End If 
						
			set objRecordset = Server.CreateObject ("ADODB.Recordset")
			set objRecordset = OBJConnection.Execute(sql)

			response.write "<BR>"
			response.write "<b>Residuos:</b><BR>"

			'Excepcion para tratar las excepciones de residuos sanitarios
			Es_Sanitario = ""
			If Len(CNAE_1) > 3 Then
				If Left(CNAE_1,3) = "851" Or Left(CNAE_1,3) = "852" Then
					Es_Sanitario = "OK"
				End If 
			End If
			If CNAE_2 <> "" Then
				If Len(CNAE_2) > 3 Then
					If Left(CNAE_2,3) = "851" Or Left(CNAE_2,3) = "852" Then
						Es_Sanitario = "OK"
					End If 
				End If
			End If
			If CNAE_3 <> "" Then
				If Len(CNAE_3) > 3 Then
					If Left(CNAE_3,3) = "851" Or Left(CNAE_3,3) = "852" Then
						Es_Sanitario = "OK"
					End If 
				End If
			End If

			if Not objRecordset.EOF Then
			
				While Not objRecordset.EOF	   
					If IsNull(objRecordset("titulo")) Then
						strTituloLey = ""
					Else
						strTituloLey = objRecordset("titulo")
					End if
					If IsNull(objRecordset("subtitulo")) Then
						strSubTituloLey = ""
					Else
						strSubTituloLey = objRecordset("subtitulo")
					End If
					If IsNull(objRecordset("texto")) Then
						strTextoLey = ""
					Else
						strTextoLey = limpiaTextoLey(objRecordset("texto"))
					End If
					If IsNull(objRecordset("idenlace")) Then
						strEnlaceLey = ""
					Else
						strEnlaceLey = CStr(objRecordset("idenlace"))
					End If			
					'Si el CNAE no es sanitario, no salen ciertas leyes de residuos sanitarios
					If CStr(objRecordset("Es_Actividad_Sanitaria")) = "0" Then
						strOK = EscribeLey(strTituloLey,strSubTituloLey,strTextoLey,strEnlaceLey)
					ElseIf Es_Sanitario = "OK" And CStr(objRecordset("Es_Actividad_Sanitaria")) = "1" Then
						strOK = EscribeLey(strTituloLey,strSubTituloLey,strTextoLey,strEnlaceLey)
					End If 
				objRecordset.MoveNext
				Wend	
			End If

			'Buscamos en Vig. legislativa aspecto ambiental Vertidos: (depende de la opcion que elija el usuario)
			If CStr(grupo_2) = "70" Or CStr(grupo_2) = "80" Or CStr(grupo_2) = "100"  Then
			
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE IdTipo_Ambiental = 2 AND ((IdAutonomia = " & idautonomia & ") OR (ambito = 1 AND (IdSubTipo_Ambiental = 7 OR IdSubTipo_Ambiental = 8))) AND Es_LegislacionOnline = 1 ORDER BY ambito"
			
			Else
			
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE IdTipo_Ambiental = 2 AND ((IdAutonomia = " & idautonomia & ") OR (ambito = 1)) AND Es_LegislacionOnline = 1 ORDER BY ambito"
					
			End If
			
				   set objRecordset = Server.CreateObject ("ADODB.Recordset")
				   set objRecordset = OBJConnection.Execute(sql)

			response.write "<BR>"
			response.write "<b>Vertidos:</b><BR>"
			
			If CStr(grupo_2) = "70" Then
				response.write "&nbsp;&nbsp;- Consultar ordenanzas municipales sobre vertidos.<BR><BR>"
			End if

			if Not objRecordset.EOF Then
			
				While Not objRecordset.EOF	   
					If IsNull(objRecordset("titulo")) Then
						strTituloLey = ""
					Else
						strTituloLey = objRecordset("titulo")
					End if
					If IsNull(objRecordset("subtitulo")) Then
						strSubTituloLey = ""
					Else
						strSubTituloLey = objRecordset("subtitulo")
					End If
					If IsNull(objRecordset("texto")) Then
						strTextoLey = ""
					Else
						strTextoLey = limpiaTextoLey(objRecordset("texto"))
					End If
					If IsNull(objRecordset("idenlace")) Then
						strEnlaceLey = ""
					Else
						strEnlaceLey = CStr(objRecordset("idenlace"))
					End If			
					strOK = EscribeLey(strTituloLey,strSubTituloLey,strTextoLey,strEnlaceLey)
				objRecordset.MoveNext
				Wend	
			End If

			'Buscamos en Vig. legislativa aspecto ambiental emisiones atmosféricas
			If CStr(grupo_3) = "110" Then
			
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE (IdTipo_Ambiental = 3 AND IdSubTipo_Ambiental = 1 AND Ambito = 1 AND Es_LegislacionOnline = 1) OR (IdTipo_Ambiental = 3 AND Ambito = 2 AND idAutonomia = " & idautonomia & " AND Es_LegislacionOnline = 1 AND idenlace <> 3004 AND idenlace <> 4066 AND idenlace <> 4063) ORDER BY ambito"
					
			End If
			
				   set objRecordset = Server.CreateObject ("ADODB.Recordset")
				   set objRecordset = OBJConnection.Execute(sql)

			response.write "<BR>"
			response.write "<b>Emisiones atmosféricas:</b><BR>"
			
			If CStr(grupo_3) = "110" Then

				if Not objRecordset.EOF Then
			
					While Not objRecordset.EOF	   
						If IsNull(objRecordset("titulo")) Then
							strTituloLey = ""
						Else
							strTituloLey = objRecordset("titulo")
						End if
						If IsNull(objRecordset("subtitulo")) Then
							strSubTituloLey = ""
						Else
							strSubTituloLey = objRecordset("subtitulo")
						End If
						If IsNull(objRecordset("texto")) Then
							strTextoLey = ""
						Else
							strTextoLey = limpiaTextoLey(objRecordset("texto"))
						End If
						If IsNull(objRecordset("idenlace")) Then
							strEnlaceLey = ""
						Else
							strEnlaceLey = CStr(objRecordset("idenlace"))
						End If			
						strOK = EscribeLey(strTituloLey,strSubTituloLey,strTextoLey,strEnlaceLey)
					objRecordset.MoveNext
					Wend	
				End If
			Else
				response.write "&nbsp;&nbsp;La actividad de la empresa no está afectada por la legislación ambiental en este aspecto.<BR>"
			End If 

			'Comprobamos COVS
			encontrado = 0
If 1= 0 Then
			for i = 0 to UBound(miTabla,2) 
				If CStr(mitabla(4,i)) = "1" Then
					If encontrado = 0 Then
						encontrado = 1
						actividad_afectada = 1
						response.write "<b>Emisión de COVS:</b>&nbsp;" & strCOVS & "<br>"
					End If
					response.write "&nbsp;&nbsp;<b>- CNAE " & CStr(mitabla(1,i)) & ".</b> " & CStr(mitabla(2,i)) & "<br>"
				End If
				Select Case CStr(mitabla(1,i))
					Case "505", "5050", "50500", "5151", "51510", "602", "6024", "60242"
						CNAE_Encontrado = "OK"
				End select
			Next
End If 
			If encontrado = 0 Then
				response.write "<BR>"
				response.write "<b>Emisión de COVS:</b><br>"
			End If
			'Empresa que usa disolventes
			If CStr(grupo_4) = "130" Or CStr(grupo_4) = "140" Then
				response.write "&nbsp;&nbsp;<A HREF='abreenlace.asp?idenlace=2926' target='_blank'>RD 117/2003 de 3 de enero, sobre limitación de emisiones de compuestos orgánicos volátiles debidas al uso de disolventes en determinadas actividades.</A>"
				response.write "<br>"
				response.write "&nbsp;&nbsp;<A HREF='abreenlace.asp?idenlace=4135' target='_blank'>Real Decreto 227/2006, de 24 de febrero, por el que se complementa el régimen jurídico sobre la limitación de las emisiones de compuestos orgánicos volátiles en determinadas pinturas y barnices y en productos de renovación del acabado de vehículos.</A>"
				response.write "<br>"
				If CStr(idautonomia) = "2" Then
					response.write "&nbsp;&nbsp;<A HREF='abreenlace.asp?idenlace=3004' target='_blank'>Decreto 231/2004, de 2 de noviembre, del Gobierno de Aragón, por el que se crea el Registro de actividades industriales emisoras de compuestos orgánicos volátiles en la Comunidad Autónoma de Aragón.</A>"
					response.write "<br>"
				ElseIf CStr(idautonomia) = "5" Then 
					response.write "&nbsp;&nbsp;<A HREF='abreenlace.asp?idenlace=4066' target='_blank'>Decreto 39/2007, de 3 de mayo, por el que se crea el Registro de Instalaciones Emisoras de Compuestos Orgánicos Volátiles (Cov) de Castilla y León</A>"
					response.write "<br>"				
				ElseIf CStr(idautonomia) = "16" Then 
					response.write "&nbsp;&nbsp;<A HREF='abreenlace.asp?idenlace=4063' target='_blank'>Decreto 19/2007, de 20 de abril, por el que se crea el registro de instalaciones que usan disolevntes orgánicos en determinadas actividades y se regula el seguimiento y control de sus emisones de compuestos orgánicos volátiles</A>"
					response.write "<br>"
				End If 
			Else 
				response.write "<tr><td style=padding-left:10px;>La actividad de la empresa no está afectada por la legislación ambiental en este aspecto."
				response.write "<br>"
			End If 
			If CNAE_Encontrado = "OK" Then
				response.write "&nbsp;&nbsp;RD 2102/1996, de 20 de septiembre, sobre el control de emisiones de compuestos orgánicos volátiles (COV) resultantes de almacenamiento y distribución de gasolina desde las terminales a las estaciones de servicio.<br>"
				response.write "&nbsp;&nbsp;R D 1437/2002, de 27 de diciembre, por el que se adecuan las cisternas de gasolina al Real Decreto 2102/1996, de 20 de septiembre, sobre control de emisiones de compuestos orgánicos volátiles (C. O. V.).<br>"
			End If 

			'Comprobamos COP
			encontrado = 0
If 1= 0 Then
			for i = 0 to UBound(miTabla,2) 
				If CStr(mitabla(5,i)) = "1" Then
					If encontrado = 0 Then
						encontrado = 1
						actividad_afectada = 1
						response.write "<BR>"
						response.write "<b>Emisión de COP:</b>&nbsp;" & strCOP & "<br>"
					End If
					response.write "&nbsp;&nbsp;<b>- CNAE " & CStr(mitabla(1,i)) & ".</b> " & CStr(mitabla(2,i)) & "<br>"
				End if
			Next
End If 

			'Comprobamos SEVESO
			encontrado = 0
If 1= 0 Then
			for i = 0 to UBound(miTabla,2) 
				If CStr(mitabla(6,i)) = "1" Then
					If encontrado = 0 Then
						encontrado = 1
						actividad_afectada = 1
						response.write "<BR>"
						response.write "<b>Acidentes mayores (Seveso):</b><br>"
					End If
					response.write "&nbsp;&nbsp;<b>- CNAE " & CStr(mitabla(1,i)) & ".</b> " & CStr(mitabla(2,i)) & "<br>"
				End if
			Next
End If 

			If encontrado = 1 Then
				'Buscamos en Vig. legislativa aspecto ambiental SEVESO
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE IdTipo_Ambiental = 6 AND (IdSubtipo_Ambiental = 22 OR IdSubtipo_Ambiental = 23) AND Es_LegislacionOnline = 1"
				
				set objRecordset = Server.CreateObject ("ADODB.Recordset")
				set objRecordset = OBJConnection.Execute(sql)

				if Not objRecordset.EOF Then
				
					While Not objRecordset.EOF	   
						If IsNull(objRecordset("titulo")) Then
							strTituloLey = ""
						Else
							strTituloLey = objRecordset("titulo")
						End if
						If IsNull(objRecordset("subtitulo")) Then
							strSubTituloLey = ""
						Else
							strSubTituloLey = objRecordset("subtitulo")
						End If
						If IsNull(objRecordset("texto")) Then
							strTextoLey = ""
						Else
							strTextoLey = limpiaTextoLey(objRecordset("texto"))
						End If
						If IsNull(objRecordset("idenlace")) Then
							strEnlaceLey = ""
						Else
							strEnlaceLey = CStr(objRecordset("idenlace"))
						End If			
						strOK = EscribeLey(strTituloLey,strSubTituloLey,strTextoLey,strEnlaceLey)
					objRecordset.MoveNext
					Wend	
				End If

			End If
			
			If es_GEI = 1 Then
				'Buscamos en Vig. legislativa aspecto ambiental GEI (gases efecto invernadero)
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE IdTipo_Ambiental = 3 AND IdSubtipo_Ambiental = 3 AND Es_LegislacionOnline = 1"
				
				set objRecordset = Server.CreateObject ("ADODB.Recordset")
				set objRecordset = OBJConnection.Execute(sql)
	
				response.write "<BR>"
				response.write "<b>Emisión de gases de efecto invernadero:</b><br>"

				if Not objRecordset.EOF Then
				
					While Not objRecordset.EOF	   
						If IsNull(objRecordset("titulo")) Then
							strTituloLey = ""
						Else
							strTituloLey = objRecordset("titulo")
						End if
						If IsNull(objRecordset("subtitulo")) Then
							strSubTituloLey = ""
						Else
							strSubTituloLey = objRecordset("subtitulo")
						End If
						If IsNull(objRecordset("texto")) Then
							strTextoLey = ""
						Else
							strTextoLey = limpiaTextoLey(objRecordset("texto"))
						End If
						If IsNull(objRecordset("idenlace")) Then
							strEnlaceLey = ""
						Else
							strEnlaceLey = CStr(objRecordset("idenlace"))
						End If			
						strOK = EscribeLey(strTituloLey,strSubTituloLey,strTextoLey,strEnlaceLey)
					objRecordset.MoveNext
					Wend	
				End If

			End If
			
			'Comprobamos contaminacion atmosferica
			encontrado = 0
If 1 = 0 then
			for i = 0 to UBound(miTabla,2) 
				If CStr(mitabla(7,i)) = "1" Then
					If encontrado = 0 Then
						encontrado = 1
						actividad_afectada = 1
						response.write "<BR>"
						response.write "<b>Contaminación atmosférica:</b>&nbsp;" & strContAtmosferica & "<br>"
					End If
					response.write "&nbsp;&nbsp;<b>- CNAE " & CStr(mitabla(1,i)) & ".</b> " & CStr(mitabla(2,i)) & "<br>"
				End if
			Next
End If

			If encontrado = 1 Then
				'Buscamos en Vig. legislativa aspecto ambiental contaminacion atmosferica
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE IdTipo_Ambiental = 3 AND IdSubtipo_Ambiental = 1 AND Es_LegislacionOnline = 1"
				
				set objRecordset = Server.CreateObject ("ADODB.Recordset")
				set objRecordset = OBJConnection.Execute(sql)

				if Not objRecordset.EOF Then
				
					While Not objRecordset.EOF	   
						If IsNull(objRecordset("titulo")) Then
							strTituloLey = ""
						Else
							strTituloLey = objRecordset("titulo")
						End if
						If IsNull(objRecordset("subtitulo")) Then
							strSubTituloLey = ""
						Else
							strSubTituloLey = objRecordset("subtitulo")
						End If
						If IsNull(objRecordset("texto")) Then
							strTextoLey = ""
						Else
							strTextoLey = limpiaTextoLey(objRecordset("texto"))
						End If
						If IsNull(objRecordset("idenlace")) Then
							strEnlaceLey = ""
						Else
							strEnlaceLey = CStr(objRecordset("idenlace"))
						End If			
						strOK = EscribeLey(strTituloLey,strSubTituloLey,strTextoLey,strEnlaceLey)
					objRecordset.MoveNext
					Wend	
				End If

			End if

			'Buscamos en Vig. legislativa aspecto ambiental cont. acustica
			sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE IdTipo_Ambiental = 4 AND (IdSubtipo_Ambiental = 10 OR IdAutonomia = " & idautonomia & ") AND Es_LegislacionOnline = 1 ORDER BY ambito"
			
			set objRecordset = Server.CreateObject ("ADODB.Recordset")
			set objRecordset = OBJConnection.Execute(sql)

			response.write "<BR>"			
			response.write "<b>Contaminación acústica:</b><br>"

			if Not objRecordset.EOF Then
			
				While Not objRecordset.EOF	   
					If IsNull(objRecordset("titulo")) Then
						strTituloLey = ""
					Else
						strTituloLey = objRecordset("titulo")
					End if
					If IsNull(objRecordset("subtitulo")) Then
						strSubTituloLey = ""
					Else
						strSubTituloLey = objRecordset("subtitulo")
					End If
					If IsNull(objRecordset("texto")) Then
						strTextoLey = ""
					Else
						strTextoLey = limpiaTextoLey(objRecordset("texto"))
					End If
					If IsNull(objRecordset("idenlace")) Then
						strEnlaceLey = ""
					Else
						strEnlaceLey = CStr(objRecordset("idenlace"))
					End If			
					strOK = EscribeLey(strTituloLey,strSubTituloLey,strTextoLey,strEnlaceLey)
				objRecordset.MoveNext
				Wend	
			End If


			'Buscamos en Vig. legislativa aspecto ambiental Sustancias y preparados peligrosos

			response.write "<BR>"
			response.write "<b>Sustancias y preparados peligrosos (accidentes graves):</b><br>"

			If CStr(grupo_5) = "160" Then

				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE IdTipo_Ambiental = 6 AND IdSubtipo_Ambiental = 25 AND Es_LegislacionOnline = 1"
				
				set objRecordset = Server.CreateObject ("ADODB.Recordset")
				set objRecordset = OBJConnection.Execute(sql)

				if Not objRecordset.EOF Then
				
					While Not objRecordset.EOF	   
						If IsNull(objRecordset("titulo")) Then
							strTituloLey = ""
						Else
							strTituloLey = objRecordset("titulo")
						End if
						If IsNull(objRecordset("subtitulo")) Then
							strSubTituloLey = ""
						Else
							strSubTituloLey = objRecordset("subtitulo")
						End If
						If IsNull(objRecordset("texto")) Then
							strTextoLey = ""
						Else
							strTextoLey = limpiaTextoLey(objRecordset("texto"))
						End If
						If IsNull(objRecordset("idenlace")) Then
							strEnlaceLey = ""
						Else
							strEnlaceLey = CStr(objRecordset("idenlace"))
						End If			
						strOK = EscribeLey(strTituloLey,strSubTituloLey,strTextoLey,strEnlaceLey)
					objRecordset.MoveNext
					Wend	
				End If
			Else
				response.write "&nbsp;&nbsp;La actividad de la empresa no está afectada por la legislación ambiental en este aspecto.<BR>"	
			End If 

			response.write "</td></tr></table>"

		'Si ningun CNAE afecta a la normativa
		'If actividad_afectada = 0 Then
		'	response.write "&nbsp;&nbsp;Las actividades de la empresa no afectan a la legislación medioambiental básica.<br>"
		'End If 
		
			sql = "SELECT * FROM ECO06_LEG_ONLINE_NOTAS ORDER BY IdNota"
				
				'response.write "<BR><a name='notas'><B><U>Anotaciones:</U></B></a><BR><BR>"
			response.write "<BR><a name='notas'><IMG SRC='imagenes/leg_online_anotaciones.gif' WIDTH='197' HEIGHT='32' BORDER='0' ALT=''></a><BR><BR>"

			response.write "<TABLE cellpadding='5' style='border:1px solid #000000;'><tr><td>"

				Set objRecordset4 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset4 = OBJConnection.Execute(sql)	
				While Not objRecordset4.EOF
					response.write "<b>Nota " & objRecordset4("IdNota") & ":</b>&nbsp;" & objRecordset4("Nota") & "<BR><BR>"
					objRecordset4.Movenext
				Wend  
			
			response.write "</td></tr></table>"
		End if
		%>
			<BR><BR>
	</div>
</div>
</body>
</html>

