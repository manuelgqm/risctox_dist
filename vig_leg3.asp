<!--#include file="web_inicio.asp"-->
<%
 	'Const adOpenKeyset = 1
	'DIM objConnection	
	'DIM objRecordset
	
	'Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	'OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	
	'----- Si es restringida y no estás identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	
	texto_busca = ""
	Tipo = "0"
	idSubtipo = "0"
	ambito_ter = "0"
	id_autonomia = "0"

	function limpia(texto)
		texto = replace(texto,"'","&#39;")
		limpia = texto
	end Function

	'numeracion = "AIBBBB"
	 numeracion = "AIBCBAB"
	
	seccion = asc(mid(numeracion,3,1))-64

	idpagina = 660	'--- página Buscador Vig. Legislativa (sólo para registrar estadísticas)
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
<title>ECOinformas: Leyes de vigilancia legislativa</title>
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
<link rel="stylesheet" type="text/css" href="estructura.css"  />
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
              				     response.write titulo&"<a href=index.asp?idpagina=571>Vigilancia legislativa</a>&nbsp;&gt;&nbsp;<a href='vig_leg_busca.asp'>Inicio</a></p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              				   end if %>
				
					<% sql = "SELECT * FROM ECO06_VIG_LEG_LEYES WHERE IdLey="&request("id")
			   		   set objRecordset = Server.CreateObject ("ADODB.Recordset")
			   		   set objRecordset = OBJConnection.Execute(sql) %>
	
	
				<table align="center" cellpadding="5" cellspacing="5" width="90%">
					<tr><td class="texto" width=40%><b>Ámbito territorial:</b>&nbsp;<br><br><ul>
					<% 
					
					If Not objRecordset.EOF Then
						If Not IsNull(objRecordset("ambito")) then
							ambito = objRecordset("ambito")
						Else
							ambito = "0"
						End If
						If Not IsNull(objRecordset("idautonomia")) then
							idautonomia = objRecordset("idautonomia")
						Else
							idautonomia = "0"
						End if
					End if
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

					if CStr(ambito) <> "0" then response.write "<li class=vineta>" & texto_ambito_te & "</li>" 
					%>
					    </ul>
						</td><td class="texto" width=60%><b>Aspectos ambientales:</b>&nbsp;<br><br><ul>
					<%

					If Not objRecordset.EOF Then
						If Not IsNull(objRecordset("idSubtipo_Ambiental")) then
							idSubtipo_Ambiental = objRecordset("idSubtipo_Ambiental")
						Else
							idSubtipo_Ambiental = "0"
						End If
						If Not IsNull(objRecordset("idTipo_Ambiental")) then
							idTipo_Ambiental = objRecordset("idTipo_Ambiental")
						Else
							idTipo_Ambiental = "0"
						End If
					End if
					If CStr(idSubtipo_Ambiental) <> "0" And CStr(ambito) <> "2" Then
						sql = "SELECT * FROM ECO06_VIG_LEG_TIPOS INNER JOIN ECO06_VIG_LEG_SUBTIPOS ON ECO06_VIG_LEG_TIPOS.IdTipo = ECO06_VIG_LEG_SUBTIPOS.IdTipo WHERE  IdSubTipo = " & CStr(idSubtipo_Ambiental)
						Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
						set objRecordset2 = OBJConnection.Execute(sql)
						If Not objRecordset2.EOF Then
							texto_ambito_am = objRecordset2("nombre_tipo") 
							texto_ambito_am = texto_ambito_am & " (" & objRecordset2("nombre_subtipo") & ")"
						End If
					Else
						sql = "SELECT * FROM ECO06_VIG_LEG_TIPOS WHERE IdTipo = " & CStr(idTipo_Ambiental)
						Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
						set objRecordset2 = OBJConnection.Execute(sql)
						If Not objRecordset2.EOF Then
							texto_ambito_am = objRecordset2("nombre_tipo") 
						End If						
					End If

					response.write "<li class=vineta>" & texto_ambito_am & "</li>"
					%>
					</ul>
					</td></tr>
					</table>

							

					<table align="center" cellpadding="5" cellspacing="5" width="90%">
					<tr><td rowspan=5 valign=top align="right"><br><img src="imagenes/ico_puntito.gif" valign="top"></td><td></td></tr>
					<tr><td colspan=2 class="texto" valign=top><b><%=objRecordset("titulo")%></b></td></tr>
					<tr><td colspan=2 class="texto" bgcolor="#DDDDDD"><%=objRecordset("subtitulo")%></td></tr>
					<tr><td colspan=2 class="texto"><%=objRecordset("texto")%></td></tr>
					<tr><td colspan=2 class="texto"><A HREF="abreenlace.asp?idenlace=<%=objRecordset("idenlace")%>" target="_blank">Enlace a la ley</A></td></tr>
					</table>
					
					
					<p align="center"><input type="button" class="boton" value="imprimir" onclick="print()"></p>
					<P class="titulo2">Buscador de normativa:</p>

					<form name="formulario" action="vig_leg_busca.asp?" method="POST">
					<table align="center" width="95%" cellpadding=0 cellspacing=5 border=0 style="background-color: #EFEFEF;" class="tabla2"><tr><td>
					<table style="background: url(imagenes/buscador.gif); background-repeat: no-repeat; background-position: top left; color: #EFEFEF;"><tr><td>
					<table width="95%" cellpadding=0 cellspacing=5 border=0>
					<tr><td class="texto" align="left" colspan=3>
Texto: &nbsp;<input type="text" name="texto_busca" value="<%=texto_busca%>" size="30" class="campo" maxlength="40"></td></tr>
					</table>

					<table width="95%" cellpadding=0 cellspacing=5 border=0>
					<tr><td class="texto" align="left" colspan=2 nowrap>
Ámbito:&nbsp;
<select class="campo" name="ambito" onchange="cambiaAut(document.formulario.idautonomia)">
<option value="0" <% If ambito = "" Or CStr(ambito) = "0" Then Response.write "selected" %>>&nbsp;</option>
<option value="1" <% If CStr(ambito) = "1" Then Response.write "selected" %>>Estatal</option>
<option value="2" <% If CStr(ambito) = "2" Then Response.write "selected" %>>Autonómico</option>
</select>&nbsp;

Comunidad autónoma: &nbsp;
		<select class="campo" name="idautonomia">
		<option value="0" <% If idautonomia = "" Or CStr(idautonomia) = "0" Or CStr(ambito) <> "2" Then Response.write "selected" %>>&nbsp;</option>
		<%		
			If CStr(ambito) = "2" Then
			
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
			End if
		%></select>
		</td></tr>
				<!-- <tr><td colspan=2>&nbsp;</td></tr> -->
<!-- 		<tr><td align=absmiddle colspan=2><input type="submit" value="BUSCAR" class="boton" onclick="document.buscador.submit()"></td></tr> -->
					</table>					
					
					
					<table width="95%" cellpadding=0 cellspacing=5 border=0>
					<tr><td class="texto" align="left" nowrap colspan="2">

Aspectos ambientales: &nbsp;
		<select class="campo" name="Tipo" onchange="cambia(document.formulario.Subtipo)">
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

		%></select></td></tr><tr>
					    	<td class="texto" align="left" colspan="2" nowrap>Aspectos ambientales secundarios:&nbsp;
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

		%></select></td></tr>
					</table>

<!-- 
		<BR><BR>

<CENTER><input type="submit" value="BUSCAR" class="boton" onclick="document.buscador.submit()"> </CENTER> -->
</td></tr></table>
					</td></tr><tr><td align=center><input type="submit" value="BUSCAR" class="boton" onclick="document.formulario.submit()"></td></tr>
					<tr><td align=center>&nbsp;</td></tr></table>
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
{ alert('Escribe el texto de la consulta antes de la consulta'); }

}
</script>