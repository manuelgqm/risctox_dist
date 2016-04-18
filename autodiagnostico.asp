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

	'----- Si es restringida y no estás identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina=971"

	CNAE_1 = request("CNAE_1")
	CNAE_2 = request("CNAE_2")
	CNAE_3 = request("CNAE_3")
	CNAE_1_Desc = request("CNAE_1_Desc")
	CNAE_2_Desc = request("CNAE_2_Desc")
	CNAE_3_Desc = request("CNAE_3_Desc")
	idautonomia = request("idautonomia")
	idprovincia = request("idprovincia")
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
	Centro_Autodiag = request("Centro_Autodiag")
	Consulta = request("Consulta")
	Nuevo = request("Nuevo")
	modificar = request("modificar")
	Ver = request("ver")

	'Guardamos los datos de la consulta para estudio estadistico
	If Nombre_Empresa <> "" Then
			If idautonomia = "" Then
				idautonomia_num = 0
			Else
				idautonomia_num = idautonomia
			End If
			If idprovincia = "" Then
				idprovincia_num = 0
			Else
				idprovincia_num = idprovincia
			End If
			If Grupo_1 = "" Then
				Grupo_1_num = 0
			Else
				Grupo_1_num = Grupo_1
			End If
			If Grupo_2 = "" Then
				Grupo_2_num = 0
			Else
				Grupo_2_num = Grupo_2
			End If
			If Grupo_3 = "" Then
				Grupo_3_num = 0
			Else
				Grupo_3_num = Grupo_3
			End If
			If Grupo_4 = "" Then
				Grupo_4_num = 0
			Else
				Grupo_4_num = Grupo_4
			End If
			If Grupo_5 = "" Then
				Grupo_5_num = 0
			Else
				Grupo_5_num = Grupo_5
			End If
			If Grupo_6 = "" Then
				Grupo_6_num = 0
			Else
				Grupo_6_num = Grupo_6
			End If

		If consulta = "" Then 

			sql = "INSERT INTO ECO06_AUTODIAG_CONSULTAS (Fecha, Usuario, CNAE_1, CNAE_2, CNAE_3, idAutonomia, idProvincia, CP, Grupo_1, Grupo_2, Grupo_3, Grupo_4, Grupo_5, Grupo_6, Centro_Autodiag, Nombre_Empresa, Direccion, Poblacion) VALUES ('" & now() & "', "&idusuario&", '"& CNAE_1 &"','"& CNAE_2 &"','"& CNAE_3 &"',"& idautonomia_num &","&idprovincia_num&",'"& CP &"',"&Grupo_1_num&", "&Grupo_2_num&", "&Grupo_3_num&", "&Grupo_4_num&","&Grupo_5_num&","&Grupo_6_num&", '"& Centro_Autodiag &"', '"& Nombre_Empresa &"', '"& Direccion &"', '"& Poblacion &"')"
			
			Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			Set objRecordset = OBJConnection.Execute(sql)

			sql = "SELECT Max(idConsulta) as Ultima FROM ECO06_AUTODIAG_CONSULTAS"
			
			Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			Set objRecordset = OBJConnection.Execute(sql)
			
			If Not objRecordset.EOF Then
				Consulta = objRecordset("Ultima")
			Else
				Consulta = 0
			End If
			
		Else
				
			sql = "UPDATE ECO06_AUTODIAG_CONSULTAS SET Fecha='" & now() & "', Usuario="&idusuario&", CNAE_1='"& CNAE_1 &"', CNAE_2='"& CNAE_2 &"', CNAE_3='"& CNAE_3 &"', idAutonomia="& idautonomia &", idProvincia="&idprovincia&", CP='"& CP &"', Grupo_1="&Grupo_1&", Grupo_2="&Grupo_2&", Grupo_3="&Grupo_3&", Grupo_4="&Grupo_4&", Grupo_5="&Grupo_5&", Grupo_6="&Grupo_6&", Centro_Autodiag='"& Centro_Autodiag &"', Nombre_Empresa='"& Nombre_Empresa &"', Direccion='"& Direccion &"', Poblacion='"& Poblacion &"' WHERE IdConsulta = " & consulta
			
			Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			Set objRecordset = OBJConnection.Execute(sql)

		End If
	
	Else
	
		If consulta <> "" Then
			sql = "SELECT ECO06_AUTODIAG_CONSULTAS.*, isnull(ECO06_LEG_ONLINE_CNAE.Titulo,'') AS 			CNAE_1_Desc, isnull(ECO06_LEG_ONLINE_CNAE_1.Titulo,'') AS CNAE_2_Desc, 			isnull(ECO06_LEG_ONLINE_CNAE_2.Titulo,'') AS CNAE_3_Desc FROM ECO06_LEG_ONLINE_CNAE RIGHT OUTER JOIN ECO06_LEG_ONLINE_CNAE ECO06_LEG_ONLINE_CNAE_1 RIGHT OUTER JOIN ECO06_AUTODIAG_CONSULTAS LEFT OUTER JOIN ECO06_LEG_ONLINE_CNAE ECO06_LEG_ONLINE_CNAE_2 ON                    			ECO06_AUTODIAG_CONSULTAS.CNAE_3 = ECO06_LEG_ONLINE_CNAE_2.Codigo ON                      		ECO06_LEG_ONLINE_CNAE_1.Codigo = ECO06_AUTODIAG_CONSULTAS.CNAE_2 ON                     		ECO06_LEG_ONLINE_CNAE.Codigo = ECO06_AUTODIAG_CONSULTAS.CNAE_1 WHERE IdConsulta = " & consulta

			Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			Set objRecordset = OBJConnection.Execute(sql)
			
			'response.write sql

			'Falta sacar las descripciones de los CNAES y ya está
			If Not objRecordset.EOF Then 
				CNAE_1 = objRecordset("CNAE_1")
				CNAE_2 = objRecordset("CNAE_2")
				CNAE_3 = objRecordset("CNAE_3")
				CNAE_1_Desc = objRecordset("CNAE_1_Desc")
				CNAE_2_Desc = objRecordset("CNAE_2_Desc")
				CNAE_3_Desc = objRecordset("CNAE_3_Desc")
				idautonomia = objRecordset("idautonomia")
				idprovincia = objRecordset("idprovincia")
				'empresas = objRecordset("empresas")
				Nombre_Empresa = objRecordset("Nombre_Empresa")
				Direccion = objRecordset("Direccion")
				Poblacion = objRecordset("Poblacion")
				CP = objRecordset("CP")
				Grupo_1 = objRecordset("Grupo_1")
				Grupo_2 = objRecordset("Grupo_2")
				Grupo_3 = objRecordset("Grupo_3")
				Grupo_4 = objRecordset("Grupo_4")
				Grupo_5 = objRecordset("Grupo_5")
				Grupo_6 = objRecordset("Grupo_6")
				Centro_Autodiag = objRecordset("Centro_Autodiag")


			End if

		End if
		
	End if

	
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

	idpagina = 972	'--- página Buscador Autodiag (sólo para registrar estadísticas)
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
%>
	function comprueba_pagina()
	{
		correcto = 1;

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
			//if (document.formulario.CNAE_1.value == "" && correcto == 1) 
			//	{ alert('Selecciona el código CNAE con la actividad principal de la empresa');
			//	correcto = 0;
			//	}
			if ((document.formulario.Grupo_1.value == "0" || document.formulario.Grupo_2.value == "0" || document.formulario.Grupo_3.value == "0" || document.formulario.Grupo_4.value == "0" || document.formulario.Grupo_5.value == "0" || document.formulario.Grupo_6.value == "0") && correcto == 1) 
				{ alert('Contesta a todas las cuestiones sobre datos laborales de la empresa');
				correcto = 0;
				}			
			if (correcto == 1)
				{ document.formulario.submit() ; }
			
	}

<%
  response.write "</script> " & vbcrlf
%>

</head>
<body>

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
			<P class="titulo2" style="padding-left:15px;">Autodiagnóstico:</p>

	<%  'Si el usuario tiene varios tests los mostramos
		
		If (Consulta = "" And Nuevo = "") Or Ver <> "" Then

			sql = "SELECT * FROM ECO06_AUTODIAG_CONSULTAS Where Usuario = " & idusuario & " ORDER BY IdConsulta"
			Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			objRecordset.Open sql,OBJConnection,adOpenKeyset
			num_resultados = objRecordset.recordcount
			i = 0

			If num_resultados > 0 Then
				%>
				<P style="padding-left:15px;">Tienes registrados los siguientes autodiagnósticos:</P>
				<% 
			
				response.write "<table align='center' width='95%' cellpadding=0 cellspacing=0 border=0>"
				response.write "<tr><td class='texto' colspan='3'><b>&nbsp;</b></td><td class='texto' align='center'><IMG SRC='imagenes/icon_organiz.gif' alt='Organización y gestión' align='absmiddle'></td><td class='texto' align='center'><IMG SRC='imagenes/icon_bidon.gif' alt='Aspectos ambientales' align='absmiddle'></td><td class='texto' align='center'><IMG SRC='imagenes/icon_aspeclabor2.gif' alt='Aspectos laborales' align='absmiddle'></td><tr>"
				response.write "<tr><td class='texto'><b>&nbsp;</b></td><td class='texto' width='30%'><b>Empresa</b></td><td class='texto' width='30%'><b>Centro</b></td><td class='texto' align='center'><b>Organización y gestión</b></td><td class='texto' align='center'><b>Aspectos ambientales</b></td><td class='texto' align='center'><b>Aspectos laborales</b></td><tr>"

				While Not objRecordset.EOF 
				
					i = i + 1
					strImagenes1 = ""
					strImagenes2 = ""
					strImagenes3 = ""

					sql = "SELECT * FROM ECO06_AUTODIAG_CUESTIONARIOS ORDER BY IdCuestionario"
					Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
					objRecordset2.Open sql,OBJConnection,adOpenKeyset
					
					While Not objRecordset2.EOF 
						'Ponemos el icono de cuestionario negro si ya lo ha hecho o gris si aun no lo ha hecho
						sql = "SELECT * FROM ECO06_AUTODIAG_VALORACIONES WHERE Consulta = " & objRecordset("idConsulta") & " AND Cuestionario = " & objRecordset2("IdCuestionario")
						Set objRecordset3 = Server.CreateObject ("ADODB.Recordset")
						objRecordset3.Open sql,OBJConnection,adOpenKeyset
						
						If Not objRecordset3.EOF Then 
							If objRecordset2("Bloque") = 1 then
								strImagenes1 = strImagenes1 & "<IMG SRC='imagenes/Cuestionario" & objRecordset2("IdCuestionario") & ".gif' alt='" & objRecordset2("NombreCuestionario") & "' align='absmiddle'>"
							ElseIf objRecordset2("Bloque") = 2 Then
								strImagenes2 = strImagenes2 & "<IMG SRC='imagenes/Cuestionario" & objRecordset2("IdCuestionario") & ".gif' alt='" & objRecordset2("NombreCuestionario") & "' align='absmiddle'>"
							Else
								strImagenes3 = strImagenes3 & "<IMG SRC='imagenes/Cuestionario" & objRecordset2("IdCuestionario") & ".gif' alt='" & objRecordset2("NombreCuestionario") & "' align='absmiddle'>"		
							End If 
						Else
							If objRecordset2("Bloque") = 1 then
								strImagenes1 = strImagenes1 & "<IMG SRC='imagenes/Cuestionario" & objRecordset2("IdCuestionario") & "g.gif' alt='" & objRecordset2("NombreCuestionario") & "' align='absmiddle'>"
							ElseIf objRecordset2("Bloque") = 2 Then
								strImagenes2 = strImagenes2 & "<IMG SRC='imagenes/Cuestionario" & objRecordset2("IdCuestionario") & "g.gif' alt='" & objRecordset2("NombreCuestionario") & "' align='absmiddle'>"
							Else
								strImagenes3 = strImagenes3 & "<IMG SRC='imagenes/Cuestionario" & objRecordset2("IdCuestionario") & "g.gif' alt='" & objRecordset2("NombreCuestionario") & "' align='absmiddle'>"			
							End If 
						End If 
						objRecordset2.MoveNext

					Wend
					
					response.write "<tr><td class='texto'><a href='autodiagnostico.asp?Consulta=" & objRecordset("idConsulta") & "'><IMG SRC='imagenes/icon_autodiagnostico.gif' alt='Realizar autodiagnóstico " & i & "' align='absmiddle'></a></td>"
					response.write "<td class='texto' title='Última modificación: " & objRecordset("Fecha") & "'><a href='autodiagnostico.asp?Consulta=" & objRecordset("idConsulta") & "&modificar=Si'>" & objRecordset("Nombre_Empresa") & "</a></td>"
					If Trim(objRecordset("Centro_Autodiag")) <> "" Then
						response.write "<td class='texto' title='Última modificación: " & objRecordset("Fecha") & "'><a href='autodiagnostico.asp?Consulta=" & objRecordset("idConsulta") & "&modificar=Si'>" & objRecordset("Centro_Autodiag") & "</a></td>" 
					else
						response.write "<td class='texto'>&nbsp;</td>" 
					End If
					'response.write "<td class='texto'>" & objRecordset("Fecha") & "</td>"
					response.write "<td class='texto' align='center'>" & strImagenes1 & "</td>"
					response.write "<td class='texto' align='center'>" & strImagenes2 & "</td>"
					response.write "<td class='texto' align='center'>" & strImagenes3 & "</td>"
					'response.write "<td class='texto'><input type='submit' value='REALIZAR " & i & "' class='boton' onclick=""location.href='autodiagnostico.asp?Consulta=" & objRecordset("idConsulta") & "'""></td>"
					response.write "</tr>"

					objRecordset.MoveNext
				Wend  
				response.write "<tr><td colspan='6' class='texto' align='left'>&nbsp;</td></tr>"
				response.write "<tr><td colspan='6' class='texto' align='left'><input type='submit' value='NUEVO' class='boton' onclick=""location.href='autodiagnostico.asp?Nuevo=Si'"">&nbsp;&nbsp;<input type='button' value='VOLVER' class='boton' onclick='javascript:history.go(-1)'></td></tr>"
				response.write "</table>"
				%>
				<P style="padding-left:15px;">Pulsa el icono <IMG SRC='imagenes/icon_autodiagnostico.gif' width='20' height='20'> de un autodiagnóstico si deseas revisar, modificar o completar tus cuestionarios. Si deseas modificar los datos de la empresa pulsa el enlace con el nombre de la empresa. <BR>Pulsa el botón “Nuevo” si deseas realizar un nuevo autodiagnóstico.</P>
				<P style="padding-left:15px;">Para cualquier duda, consulta o sugerencia sobre esta herramienta puedes dirigirte a <A HREF="mailto:autodiagnosticoambiental@ecoinformas.com?subject=Consulta enviada desde la herramienta Autodiagnóstico de ECOinformas">autodiagnosticoambiental@ecoinformas.com</A>.</P>
				<%
			Else
				%>
				<P style="padding-left:15px;">Aún no has realizado ningún test de autodiagnóstico. Pulsa el botón "Nuevo" si deseas realizar un autodiagnóstico de tu empresa.</P>
				<CENTER><input type="submit" value="NUEVO" class="boton" onclick="location.href='autodiagnostico.asp?Nuevo=Si'">&nbsp;&nbsp;<input type='button' value='VOLVER' class='boton' onclick='javascript:history.go(-1)'></CENTER>
				<BR>
				<P style="padding-left:15px;">Para cualquier duda, consulta o sugerencia sobre esta herramienta puedes dirigirte a <A HREF="mailto:autodiagnosticoambiental@ecoinformas.com?subject=Consulta enviada desde la herramienta Autodiagnóstico de ECOinformas">autodiagnosticoambiental@ecoinformas.com</A>.</P>
				<% 
			End If
			
		End If
		
		'Aqui mostrar enlaces a los distintos examenes para rellenarlos y pillar los datos de la consulta para ponerlos en el formulario
		If Consulta <> "" And modificar = "" And Ver="" Then %>
		
			<!-- <P style="padding-left:15px;">Selecciona aquí el test que quieres realizar.</P> -->
	<%		
			sql = "SELECT ECO06_AUTODIAG_BLOQUES.idBloque, ECO06_AUTODIAG_BLOQUES.Nombre_Bloque, ECO06_AUTODIAG_CUESTIONARIOS.* FROM ECO06_AUTODIAG_BLOQUES INNER JOIN ECO06_AUTODIAG_CUESTIONARIOS ON  ECO06_AUTODIAG_CUESTIONARIOS.Bloque = ECO06_AUTODIAG_BLOQUES.IdBloque ORDER BY Bloque, IdCuestionario"
			Set objRecordset = Server.CreateObject ("ADODB.Recordset")
			objRecordset.Open sql,OBJConnection,adOpenKeyset
			num_resultados = objRecordset.recordcount
			NombreBloque_old = ""
			i = 0
				%>
				<P style="padding-left:15px;">Elige cuestionario:</P>
				<% 
				While Not objRecordset.EOF 
					
					sql = "SELECT * FROM ECO06_AUTODIAG_VALORACIONES WHERE Cuestionario = " & objRecordset("idCuestionario") & " AND Consulta = " & Consulta
					Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
					objRecordset2.Open sql,OBJConnection,adOpenKeyset
					
					If Not objRecordset2.EOF Then
						strInfo = ucase(objRecordset2("Resultado"))
						strFecha = objRecordset2("Fecha_Realizacion")
					Else
						strInfo = "-"
						strFecha = "-"
					End if

					'Si el cuestionario no está aprobado, escribimos la valoración en rojo
					strClass = "texto"
					If strInfo = "DEFICIENTE" Then
						strClass = "textorojo"
					End if

					If NombreBloque_old <> objRecordset("Nombre_Bloque") Then
						If NombreBloque_old <> "" Then
							response.write "</table>"
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

					'Si el cuestionario esta hecho, ponemos el icono normal, si no ponemos el gris
					strSufijo = ""
					If strInfo = "-" Then
						strSufijo = "g"
					End if
					response.write "<tr><td class='texto' width='70%' style='padding-left:60px;'><IMG SRC='imagenes/Cuestionario" & i & strSufijo & ".gif' alt='" & objRecordset("NombreCuestionario") & "' align='absmiddle'>&nbsp;<a href='cuestionario.asp?Consulta=" & consulta & "&Bloque=" & objRecordset("Bloque") & "&Cuestionario=" & objRecordset("idCuestionario") &"'>" & objRecordset("NombreCuestionario") & "</a></td>"

					response.write "<td class='" & strClass & "' align='center'><b>" & strInfo & "</b></td>"
					response.write "<td class='texto' align='center'>" & strFecha & "</td></tr>"
					NombreBloque_old = objRecordset("Nombre_Bloque")
					objRecordset.MoveNext				
				Wend  
			response.write "</table>"

			%>
				<P><CENTER><input type="submit" value="VER INFORME COMPLETO" class="boton" onclick="location.href='Informe.asp?Consulta=<%=consulta%>'">&nbsp;&nbsp;<input type="button" value="VOLVER" class="boton" onclick="javascript:history.go(-1)"></CENTER></P>
				<P style="padding-left:15px;">Para cualquier duda, consulta o sugerencia sobre esta herramienta puedes dirigirte a <A HREF="mailto:autodiagnosticoambiental@ecoinformas.com?subject=Consulta enviada desde la herramienta Autodiagnóstico de ECOinformas">autodiagnosticoambiental@ecoinformas.com</A>.</P>
			<%
	 End If  
	 
	If (nuevo <> "") Or (modificar <> "" And consulta <> "") Then 
%>
	
			<form name="formulario" action="autodiagnostico.asp?Consulta=<%=consulta%>&Ver=Si" method="POST">
			<table align="center" width="95%" cellpadding=0 cellspacing=5 border=0 style="background-color: #EFEFEF;" class="tabla2"><tr><td>
			<table style="background: url(imagenes/buscador.gif); background-repeat: no-repeat; background-position: top left; color: #EFEFEF;"><tr><td>
			<table width="95%" cellpadding=0 cellspacing=5 border=0>
			<tr><td class="texto" colspan=2><B><U>Datos de la empresa:</U></B></td></tr>
			<tr><td class="texto" align="left" width=30% nowrap>
				Nombre empresa *:</td><td class="texto" align="left"><input type="text" name="Nombre_Empresa" value="<%=Nombre_Empresa%>" size="50" class="campo" maxlength="255"></td></tr>
			<tr><td class="texto" align="left" width=30% nowrap>
				Dirección:</td><td class="texto" align="left"><input type="text" name="Direccion" value="<%=Direccion%>" size="50" class="campo" maxlength="255"></td></tr>
			<tr><td class="texto" align="left" width=30% nowrap>
				Población:</td><td class="texto" align="left"><input type="text" name="Poblacion" value="<%=Poblacion%>" size="50" class="campo" maxlength="255"></td></tr>
			<tr><td class="texto" align="left" width=30% nowrap>
				Código postal:</td><td class="texto" align="left"><input type="text" name="CP" value="<%=CP%>" size="10" class="campo" maxlength="10"></td></tr>
			<tr><td class="texto" align="left" width=30% nowrap>
				Provincia *: </td><td class="texto" align="left"><select class="campo" name="idprovincia">
					<option value="0" <% If idprovincia = "" Or CStr(idprovincia) = "0" Then Response.write "selected" %>>-- Selecciona la provincia de tu empresa </option>
				<%	
					sql = "SELECT * FROM ECOINFORMAS_VALORES WHERE Grupo = '013' ORDER BY cast(subgrupo as int)"
		
					Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
					set objRecordset2 = OBJConnection.Execute(sql)

					While Not objRecordset2.EOF
						response.write "<option value=" & CStr(objRecordset2("subgrupo"))
						If cint(idprovincia) = cint(objRecordset2("subgrupo")) then
							response.write " selected"
						End if
						response.write ">" & CStr(objRecordset2("desc1")) & " </option>" & vbcrlf
						objRecordset2.Movenext
					wend
				%>
				</select>
				</td></tr>
			<tr><td class="texto" align="left" width=30% nowrap>
				Comunidad autónoma *: </td><td class="texto" align="left"><select class="campo" name="idautonomia">
					<option value="0" <% If idautonomia = "" Or CStr(idautonomia) = "0" Then Response.write "selected" %>>-- Selecciona la autonomía de tu empresa </option>
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
				%>
				</select>
			</td></tr>
<!-- 			<tr><td class="texto" align="left" width=30% nowrap>
				CNAE actividad principal *:</td><td class="texto" align="left">
				<input type="text" name="CNAE_1_Desc" value="<% If CNAE_1 <> "" Then response.write CNAE_1 & ". " %><%=CNAE_1_Desc%>" size="50" class="campo"></td></tr>
			<tr><td class="texto" align="left" width=30% nowrap>
				&nbsp;</td><td class="texto" align="left"><input type="hidden" name="CNAE_1" value="<%=CNAE_1%>" size="20" class="campo" maxlength="20"><input type="button" value="SELECCIONA CNAE PRINCIPAL" class="boton" onclick="window.open('listado_cnaes.asp?campo=CNAE_1','listado_CNAEs','width=640,height=200,resizable=yes,scrollbars=yes')"></td></tr>
			<tr><td class="texto" align="left" width=30% nowrap>
				CNAE actividad secundaria:</td><td class="texto" align="left">
				<input type="text" name="CNAE_2_Desc" value="<% If CNAE_2 <> "" Then response.write CNAE_2 & ". " %><%=CNAE_2_Desc%>" size="50" class="campo"></td></tr>
				<td class="texto" align="left" width=30% nowrap>
				&nbsp;</td>
				<td class="texto" align="left"><input type="hidden" name="CNAE_2" value="<%=CNAE_2%>" size="20" class="campo" maxlength="20"><input type="button" value="SELECCIONA CNAE SECUNDARIO" class="boton" onclick="window.open('listado_cnaes.asp?campo=CNAE_2','listado_CNAEs','width=640,height=200,resizable=yes,scrollbars=yes')"></td></tr>
			<tr><td class="texto" align="left" width=30% nowrap>
				CNAE otra actividad secundaria:</td><td class="texto" align="left">
				<input type="text" name="CNAE_3_Desc" value="<% If CNAE_3 <> "" Then response.write CNAE_3 & ". " %><%=CNAE_3_Desc%>" size="50" class="campo"></td></tr>
				<td class="texto" align="left" width=30% nowrap>
				&nbsp;</td>
				<td class="texto" align="left"><input type="hidden" name="CNAE_3" value="<%=CNAE_3%>" size="20" class="campo" maxlength="20"><input type="button" value="SELECCIONA OTRO CNAE SECUNDARIO" class="boton" onclick="window.open('listado_cnaes.asp?campo=CNAE_3','listado_CNAEs','width=640,height=200,resizable=yes,scrollbars=yes')"></td></tr>
 -->			<tr><td class="texto" align="left" width=30%>
				Centro para el que se realiza el autodiagnóstico:</td><td class="texto" align="left">
				<input type="text" name="Centro_Autodiag" value="<%=Centro_Autodiag%>" size="50" class="campo"></td></tr>
			<tr><td colspan=2>&nbsp;</td></tr>
			<tr><td colspan=2 class="texto"><B><U>Datos laborales:</U></B></td></tr>
			</table>
<%
				For i = 1 To 6

					Select Case i
						Case 1
							strTexto_Combo = "Número de trabajadores"
						Case 2 
							strTexto_Combo = "Número de centros de trabajo"
						Case 3
							strTexto_Combo = "Existe representación sindical"
						Case 4 
							strTexto_Combo = "Existe responsable de salud laboral"
						Case 5
							strTexto_Combo = "Existe responsable de medio ambiente"
						Case 6
							strTexto_Combo = "Existe responsable de salud laboral y medio ambiente"
					End Select 

					response.write "<table width='100%'><tr><td class='texto' align='left' width='50%'>" & strTexto_Combo & " *</td>"

					'Hacemos un bucle con las opciones de cada combo
					sql = "SELECT * FROM ECO06_AUTODIAG_COMBOS WHERE IdGrupo = " & i & " ORDER BY orden"
	
					Set objRecordset4 = Server.CreateObject ("ADODB.Recordset")
					set objRecordset4 = OBJConnection.Execute(sql)

					response.write "<td class='texto' align='left'><select class='campo' name='Grupo_" & i & "'>"
					response.write "<option value='0' selected>&nbsp;</option>"

					While Not objRecordset4.EOF
						response.write "<option value=" & CStr(objRecordset4("IdOpcion"))
						If CStr(grupo_1) = CStr(objRecordset4("IdOpcion")) Or CStr(grupo_2) = CStr(objRecordset4("IdOpcion"))  Or CStr(grupo_3) = CStr(objRecordset4("IdOpcion"))  Or CStr(grupo_4) = CStr(objRecordset4("IdOpcion"))  Or CStr(grupo_5) = CStr(objRecordset4("IdOpcion"))  Or CStr(grupo_6) = CStr(objRecordset4("IdOpcion")) then
							response.write " selected "
						End if
						response.write ">" & CStr(objRecordset4("Texto_Opcion")) & " </option>" & vbcrlf
						objRecordset4.Movenext
					wend
					response.write "</select></td></tr></table>"
				Next 
					
				%>
				
<!-- 						</td></tr>
			</table> -->
			</td></tr></table>
			</td></tr>
			<tr><td align=center>&nbsp;</td></tr>
			<tr><td align=center><input type="button" value="GUARDAR" class="boton" onclick="comprueba_pagina()">&nbsp;&nbsp;<input type="button" value="VOLVER" class="boton" onclick="javascript:history.go(-1)"></td></tr>
			<tr><td align=center>&nbsp;</td></tr>
			<tr><td align=left>(*) Campos obligatorios.</td></tr>
			</table>
<!-- <CENTER><input type="submit" value="BUSCAR" class="boton" onclick="document.buscador.submit()"> </CENTER> -->

		</form>
	<% End If %>
	

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
<script>
function enviar()
{
if (document.asesora2.consulta.value!="")
{ document.asesora2.submit(); }
else
{ alert('Escribe el texto de la consulta antes de enviar la consulta'); }

}
</script>