<!--#include file="web_inicio.asp"-->
<%
 	'Const adOpenKeyset = 1
	'DIM objConnection	
	'DIM objRecordset
	
	'Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
		
	'----- Si es restringida y no estás identificado no puedes entrar
	'if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina

	'session("id_ecogente") = "4"
	idusuario = session("id_ecogente")
	
	Consulta = request("Consulta")

	'----- Si es restringida y no estás identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina=971"

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
			response.write "<tr><td style=padding-left:10px;>- <A HREF='abreenlace.asp?idenlace=" & strEnlace & "' target='_blank'>" & strTitulo & "</A>" 
			response.write "</td></tr>"
		End If
		'''If strSubTitulo <> "" Then
			'''response.write "&nbsp;&nbsp;" & strSubTitulo & "<BR>"
		'''End If
		'''If strTexto <> "" Then
			'''response.write "&nbsp;&nbsp;" & strTexto & "<BR>"
		'''End If
		'''If strEnlace <> "" Then
			'''response.write "&nbsp;&nbsp;<A HREF='abreenlace.asp?idenlace=" & strEnlace & "' target='_blank'>Enlace a la ley</A><BR><BR>"
		'''End If
		EscribeLey = "ok"
	End Function
	
	Function FraseEmision(GRUPO_CNAE)

		frase_1 = "Autorización departamento agricultura. <a href='#notas'>Ver nota (3)</a>."
		frase_2 = "Autorización departamento obras públicas. <a href='#notas'>Ver nota (3)</a>."
		strResultado = "Autorización departamento industria. <a href='#notas'>Ver nota (3)</a>."
		
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
	

	numeracion = "AIBAB"
	seccion = asc(mid(numeracion,3,1))-64

	idpagina = 972	'--- página Buscador autodiagnostico (sólo para registrar estadísticas)
	'----- Registrar la visita
	IP = Request.ServerVariables("REMOTE_ADDR")
	Set MiBrowser = Server.CreateObject("MSWC.BrowserType")
	navegador = MiBrowser.Browser
	if session("id_ecogente")<>"" then 
		usuario = session("id_ecogente")
	else
		usuario = "4"
	end if
	orden = "INSERT INTO WEBISTAS_VISITAS (fecha,hora,IP,navegador,idpagina,idgente) VALUES ('"&date()&"','"&time()&"','"&IP&"','"&navegador&"',"&idpagina&","&usuario&")"
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	Set objRecordset = OBJConnection.Execute(orden)



	titulocompleto = ""
'	for i=2 to len(numeracion)
'		sql = "SELECT titulo,numeracion,idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='" & mid(numeracion,1,i) & "'"
'		set objRecordset = Server.CreateObject ("ADODB.Recordset")
'		set objRecordset = OBJConnection.Execute(sql)
'		if i<>2 then titulocompleto = titulocompleto & "&nbsp;&gt;&nbsp;" 
'		titulocompleto = titulocompleto & "<a href=index.asp?idpagina="&objrecordset("idpagina")&">"&objrecordset("titulo")&"</a>"
'	next 
	
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
<title>ECOinformas: Autodiagnóstico</title>
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

<%
response.write "<script language='JavaScript'>" & vbcrlf

response.write "function addOpt(oCntrl, iPos, sTxt, sVal){ " & vbcrlf
response.write "var selOpcion=new Option(sTxt, sVal); " & vbcrlf
response.write "eval(oCntrl.options[iPos]=selOpcion); " & vbcrlf
response.write "} " & vbcrlf

response.write "function cambia(oCntrl){ " & vbcrlf
response.write "while (oCntrl.length) oCntrl.remove(0); " & vbcrlf
response.write "switch (document.formulario.idautonomia.selectedIndex){ " & vbcrlf

	sql = "SELECT * FROM ECO06_LEG_ONLINE_GEI ORDER BY idAutonomia, Nombre_Empresa"
			
	Set objRecordset3 = Server.CreateObject ("ADODB.Recordset")
	set objRecordset3 = OBJConnection.Execute(sql)
	intTipo = 0
	i = -1
	While Not objRecordset3.EOF
		If CStr(intTipo) <> CStr(objRecordset3("idAutonomia")) Then
			If i > -1 Then
				response.write "break; " & vbcrlf
			End if
			response.write "case " & CStr(objRecordset3("idAutonomia")) & ":" & vbcrlf
			i = 0
			response.write "addOpt(oCntrl,  " & i & ", '', '0');" & vbcrlf 
			i = i + 1
		End If
		If Trim(objRecordset3("Nombre_Empresa")) <> "" then
			strNombre_Empresa_GEI = Replace(Trim(objRecordset3("Nombre_Empresa")),"'","´") 
		Else
			strNombre_Empresa_GEI = ""
		End If 
		response.write "addOpt(oCntrl,  " & i & ", '" & strNombre_Empresa_GEI & "', '" & CStr(objRecordset3("idInstalacion")) & "');" & vbcrlf 
		intTipo = objRecordset3("idAutonomia")
		i = i + 1 
		objRecordset3.Movenext
	Wend
	If i > -1 Then
		response.write "break; " & vbcrlf
	End if

	 response.write "}  " & vbcrlf
	response.write "}  " & vbcrlf
  response.write "</script> " & vbcrlf
%>

</head>
<body>
	
	<SCRIPT LANGUAGE="JScript">
	<!--
	function comprueba_pagina()
	{
		correcto = 1

			if (document.formulario.Nombre_Empresa.value == "" && correcto == 1) 
				{ alert('Introduce el nombre de la empresa');
				correcto = 0;
				}
			if (document.formulario.idprovincia.value == "0" && correcto == 1) 
				{ alert('Selecciona la provincia de la empresa');
				correcto = 0;
				}	
			if (document.formulario.idautonomia.value == "0" && correcto == 1) 
				{ alert('Selecciona la autonomía de la empresa');
				correcto = 0;
				}	
			if (document.formulario.CNAE_1.value == "" && correcto == 1) 
				{ alert('Selecciona el código CNAE con la actividad principal de la empresa');
				correcto = 0;
				}
			if ((document.formulario.Grupo_1.value == "0" || document.formulario.Grupo_2.value == "0" || document.formulario.Grupo_3.value == "0" || document.formulario.Grupo_4.value == "0" || document.formulario.Grupo_5.value == "0" || document.formulario.Grupo_6.value == "0") && correcto == 1) 
				{ alert('Contesta a todas las cuestiones sobre datos laborales de la empresa');
				correcto = 0;
				}			
			if (correcto == 1)
				{ document.formulario.submit() ; }
			
	}
	//-->
	</SCRIPT>

<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
			<div id="encabezado_nuevo<% response.write (seccion) %>">
			<table width="100%" cellpadding=0 border=0>
			<tr><td width="215" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="142" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="166" height="78" onclick="location.href='index.asp?idpagina=549'" style="cursor:hand">&nbsp;</td>
			    <td width="160" height="78" onclick="location.href='index.asp?idpagina=550'" style="cursor:hand">&nbsp;</td>
			    <td width="25"  height="78" align="center">
			    	<a href="mailto:salvira@istas.ccoo.es?subject=Contacto ECOinformas"><img src="imagenes/ico_contacto.gif" border="0" alt="Contacto"></a><br>
			    	<a href="busqueda.asp"><img src="imagenes/ico_busqueda.gif" border="0" alt="busqueda"></a><br>
			    	<a href="index.asp?idpagina=560"><img src="imagenes/ico_ayuda.gif" border="0" alt="ayuda"></a>
			    </td>
			</tr>
			</table>
			</div>
			<div id="menusup<% response.write (seccion) %>">
			<table border="0" cellspacing="5" cellpadding="0">
            			<tr class="textmenusup">
<%              				sql = "SELECT titulo,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion LIKE '"&mid(numeracion,1,3)&"%' AND len(numeracion)=4 ORDER BY numeracion"
			   	   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	           	set objRecordset = OBJConnection.Execute(sql)
		   	   	           	do while not objRecordset.eof
              						response.write "<td class=textmenusup>"
							if mid(numeracion,1,4)=mid(objRecordset("numeracion"),1,4) then
								response.write lcase(objRecordset("titulo"))
              						else
              							response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&" style='text-decoration:none'>"&lcase(objRecordset("titulo"))&"</a>"
              						end if
              						response.write "</td><td class=textmenusup>|</td>"
							objrecordset.movenext
 						loop %>
              			</tr>
          		</table>
			</div>
			<% if session("id_ecogente")<>"" then %>
			<div class="textsubmenu" id="submenusup<% response.write (seccion) %>">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
<%            				sql = "SELECT nombre,apellidos,sexo FROM ECOINFORMAS_GENTE WHERE idgente="&session("id_ecogente")
			   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	        set objRecordset = OBJConnection.Execute(sql)
		   	   	        usuario = objRecordset("nombre")&" "&objRecordset("apellidos")
		   	   	        usuario_sexo = "o"
		   	   	        if objRecordset("sexo")=75 then usuario_sexo = "a"
%>
            			<tr><td align="right">Usuari<%=usuario_sexo%> identificad<%=usuario_sexo%>:&nbsp;<%=usuario%>&nbsp;</td></tr>
          		</table>
			</div>
       			<% end if %>
			
			<% if 1=0 and len(numeracion)>3 then
			   sql = "SELECT titulo,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE visible=1 AND ((len(numeracion)=5 AND numeracion LIKE '"&mid(numeracion,1,4)&"%')"
			   if len(numeracion)>4 then sql = sql & " OR (len(numeracion)>4 AND numeracion LIKE '"&mid(numeracion,1,5)&"%')"
			   sql = sql & ") ORDER BY numeracion"
			   set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   set objRecordset = OBJConnection.Execute(sql)
		   	   if not objRecordset.eof then
		   	   submenu = 1 %>
		   	   <div id="margen_izquierdo<% response.write (seccion) %>">
			<% do while not objRecordset.eof %>
			<table cellpadding="5" cellspacing="1" border=0 width="95%" align="center">
			<tr>
			<% if len(objRecordset("numeracion"))=5 then %>
			<td class="campo"><img src="imagenes/flecha.gif">&nbsp;
			<% else %>
			<td class="campo" width="<%=(len(objRecordset("numeracion"))-5)*15 %>">&nbsp;</td><td class="campo" width="100%">
			<% end if %>
			<a href="index.asp?idpagina=<%=objRecordset("idpagina")%>">
			<% if objRecordset("idpagina")=idpagina then 
				'response.write "<font style='background:#EEEEEE'>"&objRecordset("titulo")&"</font>"
				response.write "<b>"&objRecordset("titulo")&"</b>"
			   else
			   	response.write objRecordset("titulo")
			   end if %>
			</a>
			</td></tr></table>
			<% objRecordset.movenext
			   loop %>
			   </div>
			<% end if
			   end if %>

			<% if submenu=1 or cstr(idpagina)="548" then %>
			<div id="interiortext">
			<% else %>
			<div id="texto">
			<% end if %>
			
				<div class="texto">
             				<% if len(numeracion)>3 then
              				     response.write "<br><p class=campo>Est&aacute;s en: "
              				     for i=1 to len(numeracion)-3
              				   	sql = "SELECT titulo,idpagina FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion='"&mid(numeracion,1,2+i)&"'" 
			   	   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	           	set objRecordset = OBJConnection.Execute(sql)
		   	   	           	response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&">"&objRecordset("titulo")&"</a>&nbsp;&gt;&nbsp;"
              				     next
              				     response.write titulo&"<a href=index.asp?idpagina=971>Autodiagnóstico</a>&nbsp;&gt;&nbsp;<a href='autodiagnostico.asp'>Inicio</a></p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              				   end if %>
					<% 
					'Construimos la cadena de filtrado
					sqlWhere = ""
					If CStr(texto_busca) <> "" Then
						sqlWhere = "WHERE (titulo like '%" & textobusqueda & "%' OR subtitulo like '%" & textobusqueda & "%' OR texto like '%" & textobusqueda & "%') "
					End If
					If CStr(Tipo)	<> "" And CStr(Tipo) <> "0" Then
						If sqlWhere = "" Then
							sqlWhere = "WHERE idTipo = " & Tipo 
						Else
							sqlWhere = sqlWhere & " AND idTipo = " & Tipo
						End if
					End If
					If CStr(idSubtipo_Ambiental) <> "" And CStr(idSubtipo_Ambiental) <> "0" Then
						If sqlWhere = "" Then
							sqlWhere = "WHERE idsubTipo_ambiental = " & idSubtipo_Ambiental 
						Else
							sqlWhere = sqlWhere & " AND idsubTipo_ambiental = " & idSubtipo_Ambiental
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

					sql = "SELECT ECO06_VIG_LEG_LEYES.*, ECO06_VIG_LEG_SUBTIPOS.* FROM ECO06_VIG_LEG_LEYES LEFT JOIN ECO06_VIG_LEG_SUBTIPOS ON idSubtipo_Ambiental = idSubtipo "

			If sqlWhere <> "" then				
					sql = sql & sqlWhere
					Set objRecordset = Server.CreateObject ("ADODB.Recordset")
					objRecordset.Open sql,OBJConnection,adOpenKeyset
					total_solicitudes = objRecordSet.recordCount
			   		num_resultados = objRecordset.recordcount
					
					If CStr(ambito) = "1" Then 
						texto_ambito_te = "Estatal"
					ElseIf CStr(ambito) = "2" Then 
						'Buscamos autonomia
						If CStr(idautonomia) <> "0" Then
								sql = "SELECT * FROM ECOINFORMAS_VALORES WHERE Grupo = '032' AND Subgrupo = '" & CStr(idautonomia) & "'"				
								Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
								set objRecordset2 = OBJConnection.Execute(sql)
								If Not objRecordset2.EOF Then
									autonomia_nombre = objRecordset2("desc1")
									texto_ambito_te = "Autonómico" & " (" & autonomia_nombre & ")"
								Else
									texto_ambito_te = "Autonómico"
								End if
						Else
							texto_ambito_te = "Autonómico"
						End if
					Else
						texto_ambito_te = ""
					End If
					
					If CStr(tipo) <> "0" Then
						sql = "SELECT * FROM ECO06_VIG_LEG_TIPOS WHERE IdTipo = '" & CStr(tipo) & "'"
						Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
						set objRecordset2 = OBJConnection.Execute(sql)
						If Not objRecordset2.EOF Then
							texto_ambito_am = objRecordset2("nombre_tipo")
						End If
						If CStr(idSubtipo_Ambiental) <> "0" Then
							sql = "SELECT * FROM ECO06_VIG_LEG_SUBTIPOS WHERE IdSubTipo = '" & CStr(idSubtipo_Ambiental) & "'"
							Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
							set objRecordset2 = OBJConnection.Execute(sql)
							If Not objRecordset2.EOF Then
								texto_ambito_am = texto_ambito_am & " (" & objRecordset2("nombre_subtipo") & ")"
							End If
						End if
					Else
						texto_ambito_am = ""
					End If
			End If
	%>
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

			%>
				<p align="center"><input type="button" value="IMPRIMIR" class="boton" onclick="window.open('imprime_autodiagnostico.asp?Consulta=<%=consulta%>','Imprime','width=660,height=580,resizable=no,scrollbars=yes,menubar=yes')">&nbsp;&nbsp;<input type="button" value="VOLVER" class="boton" onclick="javascript:history.go(-1)"></p>
				<P style="padding-left:15px;">Para cualquier duda, consulta o sugerencia sobre esta herramienta puedes dirigirte a <A HREF="mailto:autodiagnosticoambiental@ecoinformas.com?subject=Consulta enviada desde la herramienta Autodiagnóstico de ECOinformas">autodiagnosticoambiental@ecoinformas.com</A>.</P>
			<%
	 End If  
	 
	%>

					</div>
				<p>&nbsp;</p>
				
			</div>
			
			<map name="Map1" id="Map1">
            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="400,38,546,104" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="547,38,701,104" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<map name="Map2" id="Map2">
            		<area shape="rect" coords="300,8,392,66" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="393,8,539,66" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,8,694,66" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
      			<map name="Map3" id="Map3">
            		<area shape="rect" coords="300,18,392,80" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="393,18,539,80" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,18,694,80" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<img src="imagenes/pie<% response.write (seccion) %>.jpg" width="708" border="0" usemap="#Map<% response.write (seccion) %>">
     			
    			
    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>

</body>
</html>
