<!--#include file="web_inicio.asp"-->
<%
 	'Const adOpenKeyset = 1
	'DIM objConnection	
	'DIM objRecordset
	
	'Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
		
	'----- Si es restringida y no estás identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina=710"
	
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

	Parametros = "CNAE_1=" & CNAE_1
	Parametros = Parametros & "&CNAE_2=" & CNAE_2
	Parametros = Parametros & "&CNAE_3=" & CNAE_3
	Parametros = Parametros & "&idautonomia=" & idautonomia
	Parametros = Parametros & "&empresas=" & empresas
	Parametros = Parametros & "&Nombre_Empresa=" & Nombre_Empresa
	Parametros = Parametros & "&Grupo_1=" & Grupo_1
	Parametros = Parametros & "&Grupo_2=" & Grupo_2
	Parametros = Parametros & "&Grupo_3=" & Grupo_3
	Parametros = Parametros & "&Grupo_4=" & Grupo_4
	Parametros = Parametros & "&Grupo_5=" & Grupo_5
	Parametros = Parametros & "&Grupo_6=" & Grupo_6
	Parametros = Parametros & "&Tipo_Inst=" & Tipo_Inst

	'Guardamos los datos de la consulta para estudio estadistico
	empresas = request("empresas")
	Nombre_Empresa = request("Nombre_Empresa")

	If (empresas <> "" And empresas <> "0") Or Nombre_Empresa <> "" Then
			If idautonomia = "" Then
				idautonomia_num = 0
			Else
				idautonomia_num = idautonomia
			End If
			If empresas <> "" And empresas <> "0" Then
				es_GEI = 1
			Else
				es_GEI = 0
			End if
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

		sql = "INSERT INTO ECO06_LEG_ONLINE_CONSULTAS (Fecha, CNAE_1, CNAE_2, CNAE_3, idAutonomia, Es_GEI, CP, Grupo_1, Grupo_2, Grupo_3, Grupo_4, Grupo_5, Grupo_6) VALUES ('" & now() & "','"& CNAE_1 &"','"& CNAE_2 &"','"& CNAE_3 &"',"& idautonomia_num &","&es_GEI&",'"& CP &"',"&Grupo_1_num&", "&Grupo_2_num&", "&Grupo_3_num&", "&Grupo_4_num&","&Grupo_5_num&","&Grupo_6_num&")"
		
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		Set objRecordset = OBJConnection.Execute(sql)

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
	
		'Rellenamos ahora las notas para no hacerlo luego varias veces
	sql = "SELECT * FROM ECO06_LEG_ONLINE_NOTAS"
				
		Set objRecordset6 = Server.CreateObject ("ADODB.Recordset")
		set objRecordset6 = OBJConnection.Execute(sql)	

		While Not objRecordset6.EOF
			Select Case objRecordset6("IDNota")
				Case 1
					strNota_1 = objRecordset6("Nota")
				Case 2
					strNota_2 = objRecordset6("Nota")
				Case 3
					strNota_3 = objRecordset6("Nota")
				Case 4
					strNota_4 = objRecordset6("Nota")
				Case 5
					strNota_5 = objRecordset6("Nota")
			End Select 
			objRecordset6.Movenext
		Wend 


	numeracion = "AIBBA"
	seccion = asc(mid(numeracion,3,1))-64

	idpagina = 711	'--- página Buscador Leg Online (sólo para registrar estadísticas)
	'----- Registrar la visita
	IP = Request.ServerVariables("REMOTE_ADDR")
	Set MiBrowser = Server.CreateObject("MSWC.BrowserType")
	navegador = MiBrowser.Browser
	if session("id_ecogente")<>"" then 
		usuario = session("id_ecogente")
	else
		usuario = 0
	end if
	orden = "INSERT INTO WEBISTAS_VISITAS (fecha,hora,IP,navegador,idpagina,idgente) VALUES ('"&date()&"','"&time()&"','"&IP&"','"&navegador&"',"&idpagina&","&usuario&")"
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	Set objRecordset = OBJConnection.Execute(orden)


	titulocompleto = ""
	for i=2 to len(numeracion)
		'response.write  mid(numeracion,1,i) & "<BR>"
		sql = "SELECT titulo,numeracion,idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='" & mid(numeracion,1,i) & "'"
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		if i<>2 then titulocompleto = titulocompleto & "&nbsp;&gt;&nbsp;" 
		titulocompleto = titulocompleto & "<a href=index.asp?idpagina="&objrecordset("idpagina")&">"&objrecordset("titulo")&"</a>"
	next 

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
<title>ECOinformas: Legislación Online</title>
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
		response.write "addOpt(oCntrl,  " & i & ", '" & Left(Replace(trim(objRecordset3("Nombre_Empresa")),"'","´"),50) & "', '" & CStr(objRecordset3("idInstalacion")) & "');" & vbcrlf 
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

			if (document.formulario.idautonomia.value == "0" && correcto == 1) 
				{ alert('Selecciona la autonomía de la empresa');
				correcto = 0;
				}	
			if (document.formulario.Nombre_Empresa.value == "" && document.formulario.Empresas.value == 0 && correcto == 1) 
				//CAmbiar cuando se rellenen los datos de GEI
				{ alert('Selecciona la empresa del listado de empresas emisoras o introduce su nombre');
				//{ alert('Introduce el nombre de la empresa');
				correcto = 0;
				}
			if (document.formulario.Tipo_Inst.value == "" && correcto == 1) 
				{ alert('Selecciona el tipo de instalacion de la empresa.');
				correcto = 0;
				}
			if ((document.formulario.Grupo_1.value == "0" || document.formulario.Grupo_2.value == "0" || document.formulario.Grupo_3.value == "0" || document.formulario.Grupo_4.value == "0" || document.formulario.Grupo_5.value == "0" || document.formulario.Grupo_6.value == "0") && correcto == 1) 
				{ alert('Contesta a todas las cuestiones sobre aspectos ambientales de la empresa');
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
              				     response.write titulo&"<a href=index.asp?idpagina=710>Legislación on-line</a>&nbsp;&gt;&nbsp;<a href='leg_online.asp'>Inicio</a></p>"
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
			<P class="titulo2" style="padding-left:15px;">Legislación Online:</p>

	<%  'Si el usuario ha consultado mostramos los resultados
		If (empresas <> "" And empresas <> "0") Or Nombre_Empresa <> "" Then %>
			<P style="padding-left:15px;">Resultados de la consulta de legislación onLine:</P>
			<!-- <B><U>Autorizaciones:</U></B><BR><BR> -->
			<table align="center" width="95%" cellpadding=0 cellspacing=5 border=0 style="background-color: #EFEFEF;" class="tabla2">
			<tr><td>
			<IMG SRC="imagenes/leg_online_autorizacion.gif" WIDTH="146" HEIGHT="32" BORDER="0" ALT=""></td></tr>
		<%
			'Excepcion, si se elige listado de instalaciones un tipo de empresa
			'se elige idopcion 20, si no 30 (es negativo para que salga el primero)
			If CStr(Tipo_Inst) <> "" And CStr(Tipo_Inst) <> "Ninguna de la lista" Then
				intControlCont = "20"
			Else
				intControlCont = "30"
			End If

			'Excepcion: el usuario dice que no hay calderas, pero su CNAE indica que sus 
			'actividades son peligrosas para contaminación atmosferica. 
			'Actuamos como si hubiera dicho que sí ...
			grupo_3_old = Grupo_3
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
			End if
			
			'Primero ponemos las autorizaciones, segun las opciones que selecciona el usuario
			'La opcion 10 es la de urbanismo, sale simepre
			sqlWhere = "(IdOpcion = 10 or IdOpcion = '" & Grupo_1 & "' or IdOpcion = '" & Grupo_2 & "' or IdOpcion = '" & Grupo_3 & "' or IdOpcion = '" & Grupo_4 & "' or IdOpcion = '" & Grupo_5 & "' or IdOpcion = '" & Grupo_6 & "' or IdOpcion = '" & intControlCont & "')"

			sql = "SELECT * FROM ECO06_LEG_ONLINE_AUTORIZACIONES WHERE " & sqlWhere & " ORDER BY IdOpcion"

			Set objRecordset4 = Server.CreateObject ("ADODB.Recordset")
			set objRecordset4 = OBJConnection.Execute(sql)			

			While Not objRecordset4.EOF

				select Case objRecordset4("IdOpcion")
					Case 20
						strNota = strNota_1
					Case 60
						strNota = strNota_2
					'Case 110
					'	strNota = strNota_3
					Case 130
						strNota = strNota_4
					Case 140
						strNota = strNota_5
					Case Else
						strNota = ""
				End Select
				
				'Hay una excepcion para emisiones a la atmosfera
				If objRecordset4("IdOpcion") <> 110 And objRecordset4("IdOpcion") <> 130 and objRecordset4("IdOpcion") <> 140 Then
					response.write "<tr><td>"
					response.write "<b>" & objRecordset4("Aspecto_Ambiental") & ":</b></td></tr>" 
					response.write "<tr><td title='" & strNota & "' style=padding-left:10px;>" & objRecordset4("Texto_Autorizacion")
					response.write "</td></tr>"
				Else
				'	If Len(CNAE_1) = 1 Then
				'		strLeyEmision = FraseEmision(CNAE_1 & "A")
				'	else
				'		strLeyEmision = FraseEmision(Left(CNAE_1,2))
				'	End If
					'Hacemos aqui las excepciones a las autorizaciones (COVS y emisiones atmosfericas)
					If objRecordset4("IdOpcion") = 110 then
						response.write "<tr><td>"
						response.write "<b>Emisiones atmosféricas:</b></td></tr>" 
						response.write "<tr><td title='" & strNota & "' style=padding-left:10px;>" & objRecordset4("Texto_Autorizacion")
						response.write "</td></tr>"
					ElseIf objRecordset4("IdOpcion") = 130 Then
						If intControlCont = "20" then
							response.write "<tr><td>"
							response.write "<b>Emisiones COV:</b></td></tr>" 
							response.write "<tr><td title='" & strNota_4 & "' style=padding-left:10px;>La Autorización ambiental integrada deberá incluir los Valores Límite de Emisión o los sistemas de reducción de emisiones de COVs. <a href='#notas'>Ver nota (4)</a>."
							response.write "</td></tr>"						
						Else
							response.write "<tr><td>"
							response.write "<b>Emisiones COV:</b></td></tr>" 
							response.write "<tr><td title='" & strNota_5 & "' style=padding-left:10px;>Notificación para su inscripción en el Registro de instalaciones emisoras de COVs. <a href='#notas'>Ver nota (5)</a>."
							response.write "</td></tr>"							
						End If 
					ElseIf objRecordset4("IdOpcion") = 140 then
						If intControlCont = "20" then
							response.write "<tr><td>"
							response.write "<b>Emisiones COV:</b></td></tr>" 
							response.write "<tr><td title='" & strNota_4 & "' style=padding-left:10px;>La Autorización ambiental integrada deberá incluir los Valores Límite de Emisión o los sistemas de reducción de emisiones de COVs. <a href='#notas'>Ver nota (4)</a>."
							response.write "</td></tr>"					
						Else
							response.write "<tr><td>"
							response.write "<b>Emisiones COV:</b></td></tr>" 
							response.write "<tr><td title='" & strNota_5 & "' style=padding-left:10px;>Notificación para su inscripción en el Registro de instalaciones emisoras de COVs. <a href='#notas'>Ver nota (5)</a>."
							response.write "</td></tr>"							
						End If 
					End If 
				End If 
				objRecordset4.Movenext
			wend

			'Si está afectado por la GEI, se anota ahora
			If CStr(empresas) <> "" And CStr(empresas) <> "0" then

				response.write "<tr><td>"
				response.write "<b>Emisión de gases de efecto invernadero:</b></td></tr>"
				response.write "<tr><td style=padding-left:10px;>" & strGEI
				response.write "</td></tr>"

			''''''''''La tabla GEI no tiene valores, asi que no mostramos nada ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'	sql = "SELECT * FROM ECO06_LEG_ONLINE_GEI WHERE IdInstalacion = " & Empresas
			'	
			'	Set objRecordset4 = Server.CreateObject ("ADODB.Recordset")
			'	set objRecordset4 = OBJConnection.Execute(sql)	
			'
			'	If Not objRecordset4.EOF Then
			'		response.write "<tr><td>"
			'		response.write "<b>Cantidad asignada (Tn de CO2 por año):</b></td></tr>" 
			'		response.write "<tr><td style=padding-left:10px;>"
			'		response.write objRecordset4("Anyo_1") & ": " & objRecordset4("Cantidad_1") & ";&nbsp; " & objRecordset4("Anyo_2") & ": " & 'objRecordset4("Cantidad_2") & ";&nbsp; " & objRecordset4("Anyo_3") & ": " & objRecordset4("Cantidad_3") 
			'		response.write "</td></tr>"
			'	End If 
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			End If
			
			response.write "</table>"
		%>
			<BR><BR>
			<table align="center" width="95%" cellpadding=0 cellspacing=5 border=0 style="background-color: #EFEFEF;" class="tabla2">
			<tr><td>
			<!-- <B><U>Legislación básica:</U></B><BR><BR> -->
			<IMG SRC="imagenes/leg_online_legislacion.gif" WIDTH="197" HEIGHT="32" BORDER="0" ALT=""></td></tr>
		<%
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
			End if
			

			'Primero Prevención y control de la contaminación
			If CStr(Tipo_Inst) <> "" And CStr(Tipo_Inst) <> "Ninguna de la lista" Then
				'Buscamos en Vig. legislativa aspecto ambiental : AUTORIZACIONES AMBIENTALES 
				'Subaspecto ambiental: AUTORIZACIÓN AMBIENTAL INTEGRADA (tipo 7, subtipo 28)
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE IdTipo_Ambiental = 7 AND (IdSubtipo_Ambiental = 28 OR IdAutonomia = " & idautonomia & ") AND Es_LegislacionOnline = 1 ORDER BY ambito"

			   		   set objRecordset = Server.CreateObject ("ADODB.Recordset")
			   		   set objRecordset = OBJConnection.Execute(sql)

				response.write "<tr><td>"
				response.write "<b>Prevención y control integrados de la contaminación:</b></td></tr>"

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
					response.write "<tr><td style=padding-left:10px;>La actividad de la empresa no está afectada por la legislación ambiental en este aspecto." 
					response.write "</td></tr>"
				End If
			Else
				response.write "<tr><td>"
				response.write "<b>Prevención y control integrados de la contaminación:</b></td></tr>"
				response.write "<tr><td style=padding-left:10px;>La actividad de la empresa no está afectada por la legislación ambiental en este aspecto." 
				response.write "</td></tr>"
			End if

			'Buscamos en Vig. legislativa aspecto ambiental Residuos:
			
			If CStr(grupo_1) = "50" Or CStr(grupo_1) = "60" then
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE ((IdTipo_Ambiental = 1 AND IdAutonomia = " & idautonomia & ") OR (IdTipo_Ambiental = 1 AND (IdSubtipo_Ambiental = 13 OR IdSubtipo_Ambiental = 14))) AND Es_LegislacionOnline = 1 ORDER BY ambito"
			Else
				sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE idenlace = 2946 AND Es_LegislacionOnline = 1"
			End If 
			
			set objRecordset = Server.CreateObject ("ADODB.Recordset")
			set objRecordset = OBJConnection.Execute(sql)

			response.write "<tr><td>"
			response.write "<b>Residuos:</b></td></tr>"
			
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

			response.write "<tr><td>"
			response.write "<b>Vertidos:</b></td></tr>"
			
			If CStr(grupo_2) = "70" Then
				response.write "<tr><td style=padding-left:10px;>Consultar ordenanzas municipales sobre vertidos." 
				response.write "</td></tr>"
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

			response.write "<tr><td>"
			response.write "<b>Emisiones atmosféricas:</b></td></tr>"
			
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
				response.write "<tr><td style=padding-left:10px;>La actividad de la empresa no está afectada por la legislación ambiental en este aspecto."
				response.write "</td></tr>"			
			End If 

			'Comprobamos COVS
			encontrado = 0
If 1= 0 Then 
			for i = 0 to UBound(miTabla,2) 
				If CStr(mitabla(4,i)) = "1" Then
					If encontrado = 0 Then
						encontrado = 1
						actividad_afectada = 1
						response.write "<tr><td>"
						response.write "<b>Emisión de COVS:</b></td></tr>"
						response.write "<tr><td style=padding-left:10px;>" & strCOVS 
						response.write "</td></tr>"
					End If
					response.write "<tr><td style=padding-left:20px;><b>- CNAE " & CStr(mitabla(1,i)) & ".</b> " & CStr(mitabla(2,i)) 
					response.write "</td></tr>"
				End If
				Select Case CStr(mitabla(1,i))
					Case "505", "5050", "50500", "5151", "51510", "602", "6024", "60242"
						CNAE_Encontrado = "OK"
				End select
			Next
End If 
			If encontrado = 0 Then
				response.write "<tr><td>"
				response.write "<b>Emisión de COVS:</b></td></tr>"
			End If
			'Empresa que usa disolventes
			If CStr(grupo_4) = "130" Or CStr(grupo_4) = "140" Then
				response.write "<tr><td style=padding-left:10px;><A HREF='abreenlace.asp?idenlace=2926' target='_blank'>RD 117/2003 de 3 de enero, sobre limitación de emisiones de compuestos orgánicos volátiles debidas al uso de disolventes en determinadas actividades.</A>"
				response.write "</td></tr>"
				response.write "<tr><td style=padding-left:10px;><A HREF='abreenlace.asp?idenlace=4135' target='_blank'>Real Decreto 227/2006, de 24 de febrero, por el que se complementa el régimen jurídico sobre la limitación de las emisiones de compuestos orgánicos volátiles en determinadas pinturas y barnices y en productos de renovación del acabado de vehículos.</A>"
				response.write "</td></tr>"
				If CStr(idautonomia) = "2" Then
					response.write "<tr><td style=padding-left:10px;><A HREF='abreenlace.asp?idenlace=3004' target='_blank'>Decreto 231/2004, de 2 de noviembre, del Gobierno de Aragón, por el que se crea el Registro de actividades industriales emisoras de compuestos orgánicos volátiles en la Comunidad Autónoma de Aragón.</A>"
					response.write "</td></tr>"
				ElseIf CStr(idautonomia) = "5" Then 
					response.write "<tr><td style=padding-left:10px;><A HREF='abreenlace.asp?idenlace=4066' target='_blank'>Decreto 39/2007, de 3 de mayo, por el que se crea el Registro de Instalaciones Emisoras de Compuestos Orgánicos Volátiles (Cov) de Castilla y León</A>"
					response.write "</td></tr>"				
				ElseIf CStr(idautonomia) = "16" Then 
					response.write "<tr><td style=padding-left:10px;><A HREF='abreenlace.asp?idenlace=4063' target='_blank'>Decreto 19/2007, de 20 de abril, por el que se crea el registro de instalaciones que usan disolevntes orgánicos en determinadas actividades y se regula el seguimiento y control de sus emisones de compuestos orgánicos volátiles</A>"
					response.write "</td></tr>"
				End If 
			Else 
				response.write "<tr><td style=padding-left:10px;>La actividad de la empresa no está afectada por la legislación ambiental en este aspecto."
				response.write "</td></tr>"
			End If 
			If CNAE_Encontrado = "OK" Then
				response.write "<tr><td style=padding-left:10px;>RD 2102/1996, de 20 de septiembre, sobre el control de emisiones de compuestos orgánicos volátiles (COV) resultantes de almacenamiento y distribución de gasolina desde las terminales a las estaciones de servicio."
				response.write "</td></tr>"
				response.write "<tr><td style=padding-left:10px;>RD 1437/2002, de 27 de diciembre, por el que se adecuan las cisternas de gasolina al Real Decreto 2102/1996, de 20 de septiembre, sobre control de emisiones de compuestos orgánicos volátiles (C. O. V.)."
				response.write "</td></tr>"
			End If 

			'Comprobamos COP
			encontrado = 0
If 1 = 0 Then 
			for i = 0 to UBound(miTabla,2) 
				If CStr(mitabla(5,i)) = "1" Then
					If encontrado = 0 Then
						encontrado = 1
						actividad_afectada = 1
						response.write "<tr><td>"
						response.write "<b>Emisión de COP:</b></td></tr>"
						response.write "<tr><td style=padding-left:10px;>" & strCOP
						response.write "</td></tr>"
					End If
					response.write "<tr><td style=padding-left:20px;><b>- CNAE " & CStr(mitabla(1,i)) & ".</b> " & CStr(mitabla(2,i)) 
					response.write "</td></tr>"
				End if
			Next
End If 
			'Comprobamos SEVESO
			encontrado = 0
If 1=0 then
			for i = 0 to UBound(miTabla,2) 
				If CStr(mitabla(6,i)) = "1" Then
					If encontrado = 0 Then
						encontrado = 1
						actividad_afectada = 1
						response.write "<tr><td>"
						response.write "<b>Acidentes mayores (Seveso):</b></td></tr>"
					End If
					response.write "<tr><td style=padding-left:20px;><b>- CNAE " & CStr(mitabla(1,i)) & ".</b> " & CStr(mitabla(2,i)) 
					response.write "</td></tr>"
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
				
				response.write "<tr><td>"
				response.write "<b>Emisión de gases de efecto invernadero:</b></td></tr>"

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
If 1=0 Then 
			for i = 0 to UBound(miTabla,2) 
				If CStr(mitabla(7,i)) = "1" Then
					If encontrado = 0 Then
						encontrado = 1
						actividad_afectada = 1
						response.write "<tr><td>"
						response.write "<b>Contaminación atmosférica:</b></td></tr>"
						response.write "<tr><td style=padding-left:10px;>" & strContAtmosferica
						response.write "</td></tr>"
					End If
					response.write "<tr><td style=padding-left:20px;><b>- CNAE " & CStr(mitabla(1,i)) & ".</b> " & CStr(mitabla(2,i)) 
					response.write "</td></tr>"
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
			
			response.write "<tr><td>"
			response.write "<b>Contaminación acústica:</b></td></tr>"

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

			response.write "<tr><td>"
			response.write "<b>Sustancias y preparados peligrosos (accidentes graves):</b></td></tr>"

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
				response.write "<tr><td style=padding-left:10px;>La actividad de la empresa no está afectada por la legislación ambiental en este aspecto."
				response.write "</td></tr>"			
			End If 

		response.write "</table>"
		'Si ningun CNAE afecta a la normativa
		'If actividad_afectada = 0 Then
		'	response.write "&nbsp;&nbsp;Las actividades de la empresa no afectan a la legislación medioambiental básica.<br>"
		'End If 
		
		sql = "SELECT * FROM ECO06_LEG_ONLINE_NOTAS ORDER BY IdNota"

		%>
			<BR><BR>
			<table align="center" width="95%" cellpadding=0 cellspacing=5 border=0 style="background-color: #EFEFEF;" class="tabla2">
			<tr><td>
		<%

				'response.write "<BR><a name='notas'><B><U>Anotaciones:</U></B></a><BR><BR>"
				response.write "<a name='notas'><IMG SRC='imagenes/leg_online_anotaciones.gif' WIDTH='197' HEIGHT='32' BORDER='0' ALT=''></a></td></tr>"

				Set objRecordset4 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset4 = OBJConnection.Execute(sql)	
				While Not objRecordset4.EOF
					response.write "<tr><td><b>Nota " & objRecordset4("IdNota") & ":</b></td></tr>"
					response.write "<tr><td style=padding-left:10px;>" & objRecordset4("Nota") & "</td></tr>"
					objRecordset4.Movenext
				Wend  
			response.write "</table>"
		%>
			<!-- </P> -->

		<p align="center"><input type="button" value="IMPRIMIR" class="boton" onclick="window.open('imprime_leg_online.asp?<%=Parametros%>','Imprime','width=660,height=580,resizable=no,scrollbars=yes')"></p>
		<% End If
		
		'Volvemos a poner el valor que selecciono el usuario en grupo_3
		'(por si lo hemos cambiado antes)
		Grupo_3 = grupo_3_old

		%>

		
					<form name="formulario" action="leg_online.asp?" method="POST">
					<table align="center" width="95%" cellpadding=0 cellspacing=5 border=0 style="background-color: #EFEFEF;" class="tabla2"><tr><td>
					<table style="background: url(imagenes/buscador.gif); background-repeat: no-repeat; background-position: top left; color: #EFEFEF;"><tr><td>
					<table width="95%" cellpadding=0 cellspacing=5 border=0>
					<tr><td class="texto" colspan=2><B><U>Datos de la empresa:</U></B></td></tr>
					<!-- <tr><td class="texto" align="left" width=30% nowrap>
						CNAE actividad principal *:</td><td class="texto" align="left">
						<input type="text" name="CNAE_1_Desc" value="<%=CNAE_1_Desc%>" size="50" class="campo"></td></tr>
					<tr><td class="texto" align="left" width=30% nowrap>
						&nbsp;</td><td class="texto" align="left"><input type="hidden" name="CNAE_1" value="<%=CNAE_1%>" size="20" class="campo" maxlength="20"><input type="button" value="SELECCIONA CNAE PRINCIPAL" class="boton" onclick="window.open('listado_cnaes.asp?campo=CNAE_1','listado_CNAEs','width=640,height=200,resizable=yes,scrollbars=yes')"></td></tr>
					<tr><td class="texto" align="left" width=30% nowrap>
						CNAE actividad secundaria:</td><td class="texto" align="left">
						<input type="text" name="CNAE_2_Desc" value="<%=CNAE_2_Desc%>" size="50" class="campo"></td></tr>
						<td class="texto" align="left" width=30% nowrap>
						&nbsp;</td>
						<td class="texto" align="left"><input type="hidden" name="CNAE_2" value="<%=CNAE_2%>" size="20" class="campo" maxlength="20"><input type="button" value="SELECCIONA CNAE SECUNDARIO" class="boton" onclick="window.open('listado_cnaes.asp?campo=CNAE_2','listado_CNAEs','width=640,height=200,resizable=yes,scrollbars=yes')"></td></tr>
					<tr><td class="texto" align="left" width=30% nowrap>
						CNAE otra actividad secundaria:</td><td class="texto" align="left">
						<input type="text" name="CNAE_3_Desc" value="<%=CNAE_3_Desc%>" size="50" class="campo"></td></tr>
						<td class="texto" align="left" width=30% nowrap>
						&nbsp;</td>
						<td class="texto" align="left"><input type="hidden" name="CNAE_3" value="<%=CNAE_3%>" size="20" class="campo" maxlength="20"><input type="button" value="SELECCIONA OTRO CNAE SECUNDARIO" class="boton" onclick="window.open('listado_cnaes.asp?campo=CNAE_3','listado_CNAEs','width=640,height=200,resizable=yes,scrollbars=yes')"></td></tr> -->
					<tr><td class="texto" align="left" width=30% nowrap>
						Comunidad autónoma *: </td><td class="texto" align="left"><select class="campo" name="idautonomia" onchange="cambia(document.formulario.Empresas)">
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
						<tr><td class="texto" align="left" colspan=2>Comprueba si tu centro de trabajo está en el siguiente listado de emisores de gases de efecto invernadero y selecciónala (selecciona antes la autonomía correspondiente):</td></tr>
						<tr><td class="texto" align="left" width=30% nowrap>
						Empresas y centros emisores ** :</td>
						<td class="texto" align="left">
						<select class="campo" name="Empresas">
						<% If CStr(idautonomia) = "" Or CStr(idautonomia) = "0" Then %>
							<option value=0 selected>&nbsp;</option>
						<% Else 
							response.write "<option value=0 "
							If CStr(Empresas) = "" Or CStr(Empresas) = "0" Then 
								response.write " selected "
							End If 
							response.write ">&nbsp;</option>"

							sql = "SELECT * FROM ECO06_LEG_ONLINE_GEI WHERE IdAutonomia = " & idautonomia & " ORDER BY idAutonomia, Nombre_Empresa"
		
							Set objRecordset3 = Server.CreateObject ("ADODB.Recordset")
							set objRecordset3 = OBJConnection.Execute(sql)
							intTipo = 0
							i = -1
							While Not objRecordset3.EOF
								response.write "<option value=" & objRecordset3("idInstalacion") 
								If CStr(Empresas) = CStr(objRecordset3("idInstalacion")) Then 
									response.write " selected "
								End If
								response.write ">" & Left(Replace(trim(objRecordset3("Nombre_Empresa")),"'","´"),50) & "</option>"
								objRecordset3.MoveNext
							Wend
								
						End If %>
						</select>
						</td></tr>
						<tr><td class="texto" align="left" colspan=2>Si tu empresa no se encuentra en el listado anterior, introduce estos datos:</td></tr> 

						<tr><td class="texto" align="left" width=30% nowrap>
						Nombre empresa * :</td><td class="texto" align="left"><input type="text" name="Nombre_Empresa" value="<%=Nombre_Empresa%>" size="50" class="campo" maxlength="255"></td></tr>
						<tr><td class="texto" align="left" width=30% nowrap>
						Dirección:</td><td class="texto" align="left"><input type="text" name="Direccion" value="<%=Direccion%>" size="50" class="campo" maxlength="255"></td></tr>
						<tr><td class="texto" align="left" width=30% nowrap>
						Población:</td><td class="texto" align="left"><input type="text" name="Poblacion" value="<%=Poblacion%>" size="50" class="campo" maxlength="255"></td></tr>
						<tr><td class="texto" align="left" width=30% nowrap>
						Código postal:</td><td class="texto" align="left"><input type="text" name="CP" value="<%=CP%>" size="10" class="campo" maxlength="10"></td></tr>
						<tr><td class="texto" colspan=2><B><U>Responde las siguientes cuestiones sobre aspectos medioambientales:</U></B></td></tr>
						<tr><td colspan=2>
						<table width='100%'>
						<tr><td class="texto" align="left">Tipo de instalación empresa (*)</td><td class="texto" align="left" colspan="2">&nbsp;&nbsp;<input type="text" align="right" name="Tipo_Inst" value="<%=Tipo_Inst%>" size="50" class="campo"></td></tr>
						<tr><td>&nbsp;</td><td class="texto" align="left">&nbsp;&nbsp;<input type="button" value="SELECCIONA TIPO INSTALACIÓN" class="boton" onclick="window.open('listado_control_cont.asp?','TipoInstalación','width=700,height=200,resizable=yes,scrollbars=yes')"></td></tr>
						</table>
						<%
							sql = "SELECT * FROM ECO06_LEG_ONLINE_GRUPOS ORDER BY idGrupo_Opcion"
				
							Set objRecordset3 = Server.CreateObject ("ADODB.Recordset")
							set objRecordset3 = OBJConnection.Execute(sql)

							While Not objRecordset3.EOF
								'response.write "<table><tr><td class='texto' align='left'>" &  CStr(objRecordset3("idGrupo_Opcion")) & ". " & CStr(objRecordset3("Texto_Grupo")) & "</td>"

								response.write "<table width='100%'><tr><td class='texto' align='left'>" & CStr(objRecordset3("Texto_Grupo")) & " *</td>"

								'Hacemos un bucle con las opciones de cada combo
								sql = "SELECT * FROM ECO06_LEG_ONLINE_OPCIONES WHERE IdGrupo = " & objRecordset3("idGrupo_Opcion") & " ORDER BY orden"
				
								Set objRecordset4 = Server.CreateObject ("ADODB.Recordset")
								set objRecordset4 = OBJConnection.Execute(sql)

								response.write "<td class='texto' align='right'><select class='campo' name='Grupo_" & objRecordset3("idGrupo_Opcion") & "'>"
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
								objRecordset3.Movenext
							wend

						%>
						
						</td></tr>
					</table>

					</td></tr></table>
					</td></tr><tr><td align=center><input type="button" value="CONSULTAR" class="boton" onclick="comprueba_pagina();"></td></tr>
					<tr><td align=center>&nbsp;</td></tr>
					<tr><td align=left>(*) Campos obligatorios.</td></tr>
					<tr><td align=left>(**) Es obligatorio rellenar uno de los dos campos.</td></tr></table>
		<!-- <CENTER><input type="submit" value="BUSCAR" class="boton" onclick="document.buscador.submit()"> </CENTER> -->


		</form>
		
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