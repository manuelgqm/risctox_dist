<!--#include file="web_inicio.asp"-->
<%
 	'Const adOpenKeyset = 1
	'DIM objConnection	
	'DIM objRecordset
	
	'Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
		
	'----- Si es restringida y no estás identificado no puedes entrar
	'if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	
	Consulta = request("Consulta")
			
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
<title>Imprime: Autodiagnóstico</title>
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
			<BR>
			<P class='subtitulo2' align="center">INFORME FINAL AUTODIAGNÓSTICO</p>

	<% 		
		'Aqui mostrar enlaces a los distintos examenes para rellenarlos y pillar los datos de la consulta para ponerlos en el formulario
		If Consulta <> "" Then %>
		
			<!-- <P style="padding-left:15px;">Selecciona aquí el test que quieres realizar.</P> -->
	<%		
			sql = "SELECT IdConsulta, Usuario, Fecha, Nombre_Empresa, Centro_Autodiag FROM ECO06_AUTODIAG_CONSULTAS WHERE IdConsulta = " & Consulta
			Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			objRecordset.Open sql,OBJConnection,adOpenKeyset
			strNombreEmpresa = ""
			StrFecha = ""
			strReferencia = CStr(consulta)
				If Not objRecordset.EOF Then
					If Not IsNull(objRecordset("Nombre_Empresa")) And Trim(objRecordset("Nombre_Empresa")) <> "" Then
						strNombreEmpresa = Trim(objRecordset("Nombre_Empresa"))
					End If
					If Not IsNull(objRecordset("Centro_Autodiag")) And Trim(objRecordset("Centro_Autodiag")) <> "" Then
						strNombreEmpresa = strNombreEmpresa & " (Centro: " & Trim(objRecordset("Centro_Autodiag")) & ")"
					End If
					If Not IsNull(objRecordset("Fecha")) And Trim(objRecordset("Fecha")) <> "" Then
						StrFecha = Trim(objRecordset("Fecha"))
					End If
				End if
				%>
				<P style="padding-left:15px;">
					<span class='subtitulo2'>Empresa:</span><span class='texto'> <%=strNombreEmpresa%></span><BR>
					<span class='subtitulo2'>Fecha del autodiagnóstico:</span><span class='texto'> <%=StrFecha%></span><BR>
					<span class='subtitulo2'>Referencia:</span><span class='texto'> <%=strReferencia%></span>
				</P>

				<% 

			'Ponemos primero la introducción (en todos los informes)
			'sql = "SELECT * FROM ECO06_AUTODIAG_VALORACIONES_TEXTOS WHERE Puntuacion = 100"
			'Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			'objRecordset.Open sql,OBJConnection,adOpenKeyset
			
			strIntro = "El objeto último de un diagnóstico, por básico que sea, debería centrarse en servir de guía respecto de aquellos aspectos y problemas sobre los que deberíamos centrar nuestra atención de forma prioritaria a la vez que facilitar una pequeña reflexión que sirviera de punto de partida para nuestro trabajo e intervención en la empresa. En nuestro caso ese primer objetivo de guía de atención pretendemos cubrirlo con el cuestionario. Buscar la información necesaria y cumplimentar las preguntas son el primer paso de la reflexión. El segundo paso, el informe diagnóstico será un reflejo del momento en que se hizo y de la fiabilidad de la información dada.<br>Con esta herramienta tan solo podemos enunciar un escenario básico sobre el que trabajar. Por lo que el informe que obtendremos  es un documento de trabajo y no un informe finalista y cerrado."
			'If Not objRecordset.EOF Then
				'strIntro = objRecordset("texto")
			'End If 

			If strIntro <> "" Then
				response.write "<P valign='absmiddle' style='padding-left:15px;' class='subtitulo2'>Introducción</P>"
				response.write "<table align='center' width='95%' cellpadding=0 cellspacing=5 border=0>"			
				response.write "<tr><td class='texto' style='padding-left:15px;text-align:justify;'>"&strIntro&"</td></tr>"
				response.write "</table>"
			End If 


			sql = "SELECT ECO06_AUTODIAG_BLOQUES.idBloque, ECO06_AUTODIAG_BLOQUES.Nombre_Bloque, ECO06_AUTODIAG_CUESTIONARIOS.* FROM ECO06_AUTODIAG_BLOQUES INNER JOIN ECO06_AUTODIAG_CUESTIONARIOS ON  ECO06_AUTODIAG_CUESTIONARIOS.Bloque = ECO06_AUTODIAG_BLOQUES.IdBloque ORDER BY Bloque, IdCuestionario"
			Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			objRecordset.Open sql,OBJConnection,adOpenKeyset
			num_resultados = objRecordset.recordcount
			NombreBloque_old = ""
			i = 0

			sql = "SELECT ECO06_AUTODIAG_VALORACIONES.Consulta, ECO06_AUTODIAG_VALORACIONES.Cuestionario, ECO06_AUTODIAG_VALORACIONES.Puntuacion_Cuestionario, ECO06_AUTODIAG_VALORACIONES_TEXTOS.Texto FROM ECO06_AUTODIAG_VALORACIONES INNER JOIN ECO06_AUTODIAG_VALORACIONES_TEXTOS ON ECO06_AUTODIAG_VALORACIONES.Puntuacion_Cuestionario = ECO06_AUTODIAG_VALORACIONES_TEXTOS.Puntuacion WHERE Consulta = " & Consulta & " ORDER BY ECO06_AUTODIAG_VALORACIONES.Cuestionario"

			Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
			objRecordset2.Open sql,OBJConnection,adOpenKeyset

				While Not objRecordset.EOF 

					'while Not objRecordset2.EOF Then
						strInfo = ""
						If Not objRecordset2.EOF Then
							If objRecordset("IdCuestionario") = objRecordset2("Cuestionario") then
								strInfo = objRecordset2("Texto")
							End If
						End if						

						'Si el cuestionario no está aprobado, escribimos la valoración en rojo
						strClass = "texto"
						'If strInfo = "DEFICIENTE" Then
						'	strClass = "textorojo"
						'End if

						If NombreBloque_old <> objRecordset("Nombre_Bloque") Then
							If NombreBloque_old <> "" Then
								response.write "</table>"
								i = 0
							End If 
							If objRecordset("idBloque") = 1 Then
								strIcono = "icon_organiz.gif"
							ElseIf objRecordset("idBloque") = 2 Then
								strIcono = "icon_bidon.gif"
							ElseIf objRecordset("idBloque") = 3 Then
								strIcono = "icon_aspeclabor2.gif"
							End If 
							response.write "<P valign='absmiddle' style='padding-left:15px;' class='subtitulo2'><IMG SRC='imagenes/" & strIcono & "' align='absmiddle'>&nbsp;" & objRecordset("Nombre_Bloque") & "</P>"
							response.write "<table align='center' width='95%' cellpadding=0 cellspacing=5 border=0>"
						End If
						i = i + 1

						response.write "<tr><td class='texto' style='padding-left:15px;text-align:justify;'><b>"&i&". "&objRecordset("NombreCuestionario")&"</b></td></tr>"
						response.write "<tr><td class='texto' style='padding-left:30px;text-align:justify;'>"&strInfo&"</td></tr>"
						
						If strInfo <> "" Then
							objRecordset2.MoveNext
						End if
					'Wend 
					NombreBloque_old = objRecordset("Nombre_Bloque")
					objRecordset.MoveNext				
				Wend  
			response.write "</table>"

	 End If  
	 
	%>

			<BR><BR>
	</div>
</div>
</body>
</html>

