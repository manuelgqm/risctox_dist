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

	Consulta = request("Consulta")
	Bloque = request("Bloque")
	Cuestionario = request("Cuestionario")


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

	Function sustituyeAyuda(texto)
		texto = Replace(texto,"<idAyu=","<a onclick=window.open('ver_definicion.asp?id=")
		texto = Replace(texto,"/da>","','def','width=300,height=300,scrollbars=yes,resizable=yes') style='cursor:hand'>")
		texto = Replace(texto,"</idAyuda>","</a>")
		sustituyeAyuda = texto
	End Function 

	Function DevuelveValor(id_Pregunta, id_Opcion, ValorAlternativo)
		str_Resultado = "Neutra"
		Select Case CStr(id_Pregunta)
			Case "36"
				sql = "SELECT * FROM ECO06_AUTODIAG_RESULTADOS WHERE Consulta = "&consulta&" and (pregunta_detalle = 29 or pregunta_detalle = 30)"
				
				Set objRecordset44 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset44 = OBJConnection.Execute(sql)

				While Not objRecordset44.EOF
					If objRecordset44("Valor") = "Neg" Then
						str_Resultado = ValorAlternativo
					End If
					objRecordset44.MoveNext
				Wend 

			Case "50","52"
				sql = "SELECT * FROM ECO06_AUTODIAG_RESULTADOS WHERE Consulta = "&consulta&" and pregunta_detalle = 6"
				
				Set objRecordset44 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset44 = OBJConnection.Execute(sql)

				While Not objRecordset44.EOF
					If CStr(objRecordset44("Opcion")) = "2" Or CStr(objRecordset44("Opcion")) = "3" Then
						str_Resultado = ValorAlternativo
					End If
					objRecordset44.MoveNext
				Wend 

			Case "51","53"
				sql = "SELECT * FROM ECO06_AUTODIAG_RESULTADOS WHERE Consulta = "&consulta&" and pregunta_detalle = 6"
				
				Set objRecordset44 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset44 = OBJConnection.Execute(sql)

				str_Resultado = ValorAlternativo
				While Not objRecordset44.EOF
					If CStr(objRecordset44("Opcion")) = "1" Then
						str_Resultado = "Neutra"
					End If
					objRecordset44.MoveNext
				Wend 

			Case "67"

				sql = "SELECT * FROM ECO06_AUTODIAG_RESULTADOS WHERE Consulta = "&consulta&" and pregunta_detalle = 7"
				
				Set objRecordset44 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset44 = OBJConnection.Execute(sql)

				While Not objRecordset44.EOF
					If (CStr(objRecordset44("Opcion")) = "1" And CStr(id_Opcion) = "1") Or (CStr(objRecordset44("Opcion")) = "3" And CStr(id_Opcion) = "2") Or (CStr(objRecordset44("Opcion")) = "3" And CStr(id_Opcion) = "3") Or (CStr(objRecordset44("Opcion")) = "2" And CStr(id_Opcion) = "4") Then
						str_Resultado = ValorAlternativo
					End If
					objRecordset44.MoveNext
				Wend 

			Case "80"

				sql = "SELECT * FROM ECO06_AUTODIAG_RESULTADOS WHERE Consulta = "&consulta&" and pregunta_detalle = 81"
				
				Set objRecordset44 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset44 = OBJConnection.Execute(sql)

				While Not objRecordset44.EOF
					If objRecordset44("Valor") = "Pos" Then
						str_Resultado = ValorAlternativo
					End If
					objRecordset44.MoveNext
				Wend  

			Case "117", "118", "120"

				sql = "SELECT * FROM ECO06_AUTODIAG_CONSULTAS WHERE idConsulta = "&consulta&" and Grupo_1 = 1"
				
				Set objRecordset44 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset44 = OBJConnection.Execute(sql)

				str_Resultado = ValorAlternativo
				While Not objRecordset44.EOF
					If CStr(objRecordset44("Grupo_1")) = "1" Then
						str_Resultado = "Neutra"
					End If
					objRecordset44.MoveNext
				Wend  

		End select

		DevuelveValor = str_Resultado
	End Function 

	Function DevuelveVisible(id_Pregunta, ValorBusca)
			str_Resultado = "hidden"

				sql = "SELECT * FROM ECO06_AUTODIAG_RESULTADOS WHERE Consulta = "&consulta&" and pregunta_detalle = " & id_Pregunta
				
				Set objRecordset66 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset66 = OBJConnection.Execute(sql)

				While Not objRecordset66.EOF
					If InStr(ValorBusca,CStr(objRecordset66("opcion"))) > 0 Then
						str_Resultado = "visible"
					End If
					objRecordset66.MoveNext
				Wend  
			

			DevuelveVisible = str_Resultado
	End Function 

	Function DevuelveOncheck(str_pregunta_detalle,str_IdOpcion)
		str_Resultado = ""

			sql = "SELECT ECO06_AUTODIAG_PREGUNTAS.*, ECO06_AUTODIAG_PREGUNTAS_DETALLE.* FROM ECO06_AUTODIAG_PREGUNTAS INNER JOIN ECO06_AUTODIAG_PREGUNTAS_DETALLE ON IdPregunta = Pregunta WHERE Cuestionario = "&cuestionario&" and Condicion_Pregunta = " & str_pregunta_detalle 
			
			'& " AND Condicion_Respuesta LIKE '%" & str_IdOpcion & "%'"
			
			Set objRecordset88 = Server.CreateObject ("ADODB.Recordset")
			set objRecordset88 = OBJConnection.Execute(sql)
			
			If Not objRecordset88.EOF Then
				str_Resultado = "onclick="
			End if
			
			While Not objRecordset88.EOF
				If InStr(CStr(objRecordset88("Condicion_Respuesta")),str_IdOpcion) > 0 Then
					str_Resultado = str_Resultado & "cambia('capa" & CStr(objRecordset88("idPregunta_Detalle")) & "','visible');"
				Else
					str_Resultado = str_Resultado & "cambia('capa" & CStr(objRecordset88("idPregunta_Detalle")) & "','hidden');"
				End If
				objRecordset88.MoveNext
			Wend  
			
			'If strResultado <> "" Then
			'	strResultado = strResultado
			'End If
		'response.write str_Resultado & "<BR>"
		DevuelveOncheck = str_Resultado
	End Function 


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

	If request("Respuesta") <> "" Then 
		
		sql = "SELECT ECO06_AUTODIAG_PREGUNTAS.*, ECO06_AUTODIAG_PREGUNTAS_DETALLE.*, ECO06_AUTODIAG_OPCIONES.* FROM ECO06_AUTODIAG_PREGUNTAS INNER JOIN                    ECO06_AUTODIAG_PREGUNTAS_DETALLE ON ECO06_AUTODIAG_PREGUNTAS.IdPregunta = ECO06_AUTODIAG_PREGUNTAS_DETALLE.Pregunta INNER JOIN  ECO06_AUTODIAG_OPCIONES ON ECO06_AUTODIAG_PREGUNTAS_DETALLE.Tipo_Pregunta = ECO06_AUTODIAG_OPCIONES.idTipo_Pregunta WHERE (ECO06_AUTODIAG_PREGUNTAS.Cuestionario = " & Cuestionario & ") ORDER BY ECO06_AUTODIAG_PREGUNTAS.Orden_Pregunta, ECO06_AUTODIAG_PREGUNTAS_DETALLE.Orden_Detalle, ECO06_AUTODIAG_OPCIONES.idOpcion"
				
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		idPregunta_old = 0

		'Antes de empezar, limpiamos las respuestas del cuestionario (por si lo hacen más de una vez)
		sql = "DELETE FROM ECO06_AUTODIAG_RESPUESTAS WHERE Consulta = "&consulta&" AND Cuestionario = "&cuestionario
		Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		Set objRecordset2 = OBJConnection.Execute(sql)

		While Not objRecordset.EOF 
			opcion = ""
			valor = ""
			'Para radio button sólo pasamos una vez
			If idPregunta_old <> objRecordset("idPregunta_Detalle") Then 
				If objRecordset("TextBox") <> 1 And objRecordset("TextBox") <> 2 then 
					opcion = "rad_" & objRecordset("idPregunta_Detalle")
					valor = CStr(request(opcion))
					'response.write opcion & ": " & valor & "<BR>"
				End If 
			End If 
			'Para checkbox y txtbox tenemos que recoger cada opcion
			If objRecordset("TextBox") = 1 then
				opcion = "txt_" & objRecordset("idPregunta_Detalle") & "_" & objRecordset("IdOpcion")
				valor = CStr(request(opcion))
			ElseIf objRecordset("TextBox") = 2 then
				opcion = "chk_" & objRecordset("idPregunta_Detalle") & "_" & objRecordset("IdOpcion")
				valor = CStr(request(opcion))
				If valor = "on" Then
					valor = cstr(objRecordset("IdOpcion"))
				Else
					valor = ""
				End if
			End If

			'Guardamos solo cuando hay datos
			If valor <> "" then		
				sql = "INSERT INTO ECO06_AUTODIAG_RESPUESTAS (Consulta, Cuestionario, Pregunta_Detalle, Opcion, Valor, TipoElemento) VALUES ("&consulta&","&cuestionario&"," & objRecordset("idPregunta_Detalle") & "," & objRecordset("IdOpcion") & ", '" & valor & "', " & objRecordset("TextBox") & ")"
				Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
				Set objRecordset2 = OBJConnection.Execute(sql)
			End If 

			idPregunta_old = objRecordset("idPregunta_Detalle")
			objRecordset.MoveNext
		Wend 

		'Comprobamos resultado y contamos respuestas (ECO06_AUTODIAG_RESULTADOS es una vista)
		sql = "SELECT * FROM ECO06_AUTODIAG_RESULTADOS WHERE (Cuestionario = " & Cuestionario & " AND Consulta = " & consulta & ") ORDER BY Pregunta_Detalle"
				
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)

		intPos = 0
		intNeg = 0
		intNeutra = 0
		intNula = 0
		Incremento = 0
		Incremento_Mejora = 0
		intPos_Total = 0
		intNeg_Total = 0
		While Not objRecordset.EOF
			strResultado = Trim(objRecordset("Valor"))
			If strResultado = "Pos" Then
				intPos = intPos + 1
			ElseIf strResultado = "Neg" Then
				intNeg = intNeg + 1
			ElseIf strResultado = "Neutra" Then
				intNeutra = intNeutra + 1
			ElseIf strResultado = "Nula" Then
				intNula = intNula + 1
			ElseIf strResultado = "Mejora" Then
				intPos = intPos + 1
				If objRecordset("Incremento") > Incremento_Mejora Then
					Incremento_Mejora = objRecordset("Incremento")
				End if
			ElseIf strResultado = "Neg Total" Then
				intNeg = intNeg + 1
				intNeg_Total = intNeg_Total + 1
			ElseIf strResultado = "Pos Total" Then
				intPos = intPos + 1
				intPos_Total = intPos_Total + 1
			ElseIf strResultado = "Neutra/Neg" Then
				'Crear funcion que determine si es positivo o negativo en funcion de otras respuestas ...
				strResultado = DevuelveValor(objRecordset("Pregunta_Detalle"),objRecordset("Opcion"),"Neg")
				If strResultado = "Neutra" Then
					intNeutra = intNeutra + 1
				Else
					intNeg = intNeg + 1
				End if
			ElseIf strResultado = "Neutra/Pos" Then
				'Crear funcion que determine si es positivo o negativo en funcion de otras respuestas ...
				strResultado = DevuelveValor(objRecordset("Pregunta_Detalle"),objRecordset("Opcion"),"Pos")
				If strResultado = "Neutra" Then
					intNeutra = intNeutra + 1
				Else
					intPos = intPos + 1
				End if
			End If 
			If strResultado <> "Mejora" Then
				If objRecordset("Incremento") > Incremento Then
					Incremento = objRecordset("Incremento")
				End If
			End if
			objRecordset.MoveNext
		wend

		'Comprobamos si ya ha hecho este cuestionario y lo actualizamos o insertamos nuevo
		sql = "SELECT * FROM ECO06_AUTODIAG_VALORACIONES WHERE (Cuestionario = " & Cuestionario & " AND Consulta = " & consulta & ")"
		
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)

		If objRecordset.EOF then
			sql = "INSERT INTO ECO06_AUTODIAG_VALORACIONES (Consulta, Cuestionario, Fecha_Realizacion, Respuestas_Positivas, Respuestas_Negativas, Respuestas_Neutras, Respuestas_Nulas, Incremento) VALUES ("&consulta&", "&cuestionario&", '"&Date()&"', "&intPos&", "&intNeg&", "&intNeutra&", "&intNula&", "&Incremento&")"
		Else
			sql = "UPDATE ECO06_AUTODIAG_VALORACIONES SET Fecha_Realizacion = '"&Date()&"', Respuestas_Positivas = "&intPos&", Respuestas_Negativas = "&intNeg&", Respuestas_Neutras = "&intNeutra&", Respuestas_Nulas = "&intNula&", Incremento = "&Incremento&" WHERE Consulta = "&consulta&" AND Cuestionario = "&cuestionario
		End If
		
		Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		set objRecordset2 = OBJConnection.Execute(sql)

	End If 

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

	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function cambia(capa,valor) {

	if (valor == 'hidden')
		{
			document.getElementById(capa).className = "invisible_tab";
		}
	else
		{
			document.getElementById(capa).className = "visible_tab";
		}
	}

	function comprobar() 
	{
		resultado=1;
		<% for j = 1 to UltimaId %>
			if (capa<%=j%>.style.visibility == 'visible' && (document.formulario.txt_<%=j%>.value == '' || document.formulario.txt_<%=j%>.value == '0' ))
			{
				resultado = 0;
				alert('Debe rellenar la información adicional de las frases R que lo requieran.');
				document.formulario.txt_<%=j%>.focus();
			}
		<% next %> 

		if (resultado == 1)
		{
			document.formulario.submit();
		}
	}

	function comprueba_pagina()
		 {
			correcto = 1;
			radio = 1;
			nombre = '';
			nombre_old = '';
			//alert('Hay ' + document.formulario.elements.length + ' campos en el formulario');
			//nombre_capa = nombre.substring(5,nombre.length);
		for (i=0; ele=document.formulario.elements[i] ;i++ )
		{
			//alert(nombre + ': ' + ele.name + ':' + ele.style.visbility);
			if (ele.type=='text')
			{
				if (document.getElementById('capa20').className == 'visible_tab')
				{
					if(ele.type=='text' && ele.value=='' && correcto == 1) {
					  alert('Hay campos en blanco');
					  correcto = 0;
					} 
				} else {
					ele.value = '';
					//correcto = 1;
				}
			}

			if (ele.type=='radio' && correcto == 1)
			{
				nombre = ele.name;
				if (nombre != nombre_old)
				{
					nombre_capa = nombre.substring(4,nombre.length);
					capa = 'capa'+nombre_capa;
					//alert(capa);
					//if (document.getElementById(capa).style.visibility == 'visible')
					//{
					//	alert(capa+' visible');
					//} else {
					//	alert(capa+' hidden');
					//}
					//alert(nombre);
					//alert(nombre_capa+'_capa');
					if (radio == 0)
					{
						alert('Debes contestar todas las preguntas');
						correcto = 0;
					}
					radio = 0;
				}
				if (document.getElementById(capa).className == 'visible_tab')
				{
					if (ele.checked)
					{
						radio = 1;
					}
				} else {
					radio = 1;
					ele.checked = false;
				}
				nombre_old = nombre;				
			}
		}
		
		 if (nombre_old != '' && radio == 0 && correcto == 1)
			  {
					alert('Debes contestar todas las preguntas');
					correcto = 0;
			  }
		
		  
			if (correcto == 1)
				{ document.formulario.submit() ; }
		 

	}
						
	//-->
	</SCRIPT>

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

	<%  
		If request("Respuesta") <> "" Then
		
				sql = "SELECT ECO06_AUTODIAG_BLOQUES.idBloque, ECO06_AUTODIAG_BLOQUES.Nombre_Bloque, ECO06_AUTODIAG_CUESTIONARIOS.IdCuestionario, ECO06_AUTODIAG_CUESTIONARIOS.NombreCuestionario FROM         ECO06_AUTODIAG_BLOQUES INNER JOIN ECO06_AUTODIAG_CUESTIONARIOS ON ECO06_AUTODIAG_BLOQUES.idBloque = ECO06_AUTODIAG_CUESTIONARIOS.Bloque WHERE IdCuestionario = " & Cuestionario
				Set objRecordset = Server.CreateObject ("ADODB.Recordset")
				set objRecordset = OBJConnection.Execute(sql)
				
				strNombre_Bloque = ""
				strCuestionario = ""
				strImagen = ""
				strIcono = ""
				If Not objRecordset.EOF Then
					If objRecordset("idBloque") = 1 Then
						strIcono = "icon_organiz.gif"
					ElseIf objRecordset("idBloque") = 2 Then
						strIcono = "icon_bidon.gif"
					ElseIf objRecordset("idBloque") = 3 Then
						strIcono = "icon_aspeclabor2.gif"
					End If 
					strImagen = "<IMG SRC='imagenes/" & strIcono & "' align='absmiddle'>"
					strNombre_Bloque = objRecordset("Nombre_Bloque")
					strCuestionario = objRecordset("NombreCuestionario")
				End if

	%>

			<table width="95%" cellpadding=0 cellspacing=5 border=0>
			<tr><td class="texto" colspan=2><B><U>Resultados del cuestionario <%=strCuestionario%>:</U></B></td></tr>
			<tr><td class="texto" colspan=2>&nbsp;</td></tr>
			<tr><td class="texto" colspan=2 style="padding-left:15px;">Valoraciones correctas:&nbsp;<%=intPos%></td></tr>
			<tr><td class="texto" colspan=2 style="padding-left:15px;">Valoraciones incorrectas:&nbsp;<%=intNeg%></td></tr>
			<tr><td class="texto" colspan=2 style="padding-left:15px;">Respuestas no aplicables:&nbsp;<%=intNeutra%></td></tr>
			<tr><td class="texto" colspan=2 style="padding-left:15px;">Respuestas lo desconozco:&nbsp;<%=intNula%></td></tr>
			<tr><td class="texto" colspan=2>&nbsp;</td></tr>
	<%

			strNota = "No valorable"
			intNota = 0
			intPuntuacion = 0
			
			If (intPos+intNeg+intNeutra+intNula) > 0 then
				If intNeg_Total > 0 Then
					strNota = "Deficiente"
					'Hallamos id puntuacion para los textos
					sql = "SELECT * FROM ECO06_AUTODIAG_EVALUACION WHERE Cuestionario = "&cuestionario&" AND Valoracion_Final = 'Deficiente'"
					
					Set objRecordset = Server.CreateObject ("ADODB.Recordset")
					set objRecordset = OBJConnection.Execute(sql)

					If Not objRecordset.EOF Then
						intPuntuacion = objRecordset("idPuntuacion")
					Else
						intPuntuacion = 0
					End if
				%>
							<tr><td class="texto" colspan=2 style="padding-left:15px;">La valoración final del cuestionario es negativa, debido a que una o varias de las respuestas invalida toda la situación de la empresa.</td></tr>
				<%
				
				elseIf intPos_Total > 0 Then
					strNota = "Eficiente"
					'Hallamos id puntuacion para los textos
					sql = "SELECT * FROM ECO06_AUTODIAG_EVALUACION WHERE Cuestionario = "&cuestionario&" AND Valoracion_Final = 'Eficiente'"
					
					Set objRecordset = Server.CreateObject ("ADODB.Recordset")
					set objRecordset = OBJConnection.Execute(sql)					

					If Not objRecordset.EOF Then
						intPuntuacion = objRecordset("idPuntuacion")
					Else
						intPuntuacion = 0
					End if
				%>
							<tr><td class="texto" colspan=2 style="padding-left:15px;">La valoración final del cuestionario es positiva, debido a que una o varias de las respuestas valida toda la situación de la empresa.</td></tr>
				<%
				
				Else 
					If (intNula/(intPos+intNeg+intNeutra+intNula)) * 100 > 30.0 Then
						strNota = "No valorable"
						intPuntuacion = 0
			%>
						<tr><td class="texto" colspan=2 style="padding-left:15px;">Debido al gran número de respuestas "lo desconozco" de este cuestionario no es posible dar una valoración.</td></tr>
			<%
					Else 
						'Calculamos el grupo al que pertenece por nº de empleados
						If intPos+intNeg > 0 then
							intNota = (cint((intPos/(intPos+intNeg)) * 100)) + Incremento_Mejora
						Else
							intNota = 0
						End If
						If intNota > 100 Then
							intNota = 100
						End If 
						sql = "SELECT ECO06_AUTODIAG_CONSULTAS.IdConsulta, ECO06_AUTODIAG_COMBOS.IdGrupo, ECO06_AUTODIAG_COMBOS.Orden, ECO06_AUTODIAG_COMBOS.Texto_Opcion FROM         ECO06_AUTODIAG_CONSULTAS INNER JOIN ECO06_AUTODIAG_COMBOS ON ECO06_AUTODIAG_CONSULTAS.Grupo_1 = ECO06_AUTODIAG_COMBOS.Orden WHERE (ECO06_AUTODIAG_CONSULTAS.IdConsulta = "&consulta&") AND (ECO06_AUTODIAG_COMBOS.IdGrupo = 1)"
						
						Set objRecordset = Server.CreateObject ("ADODB.Recordset")
						set objRecordset = OBJConnection.Execute(sql)
						
						If Not objRecordset.EOF Then
							strGrupoValoracion = objRecordset("Orden")
							'Buscaremos los resultados en tablas y los textos a mostrar
							sql = "SELECT * FROM ECO06_AUTODIAG_EVALUACION WHERE Cuestionario = "&cuestionario&" ORDER BY idPuntuacion"
							
							Set objRecordset = Server.CreateObject ("ADODB.Recordset")
							set objRecordset = OBJConnection.Execute(sql)
							
							While Not objRecordset.EOF And intPuntuacion = 0
								intPuntuacion = 0
								incremento = 0
								'Para evitar pasarnos de 100 en el límite superior o inferior
								If objRecordset("limite_Inf_" & strGrupoValoracion) + incremento > 100 Then
									incremento = objRecordset("limite_Inf_" & strGrupoValoracion) + incremento - 101
								ElseIf objRecordset("limite_Sup_" & strGrupoValoracion) + incremento > 100 Then
									incremento = objRecordset("limite_Sup_" & strGrupoValoracion) + incremento - 101
								End if
								If objRecordset("limite_Inf_" & strGrupoValoracion) + incremento <= intNota And objRecordset("limite_Sup_" & strGrupoValoracion) + incremento >= intNota Then
									strNota = objRecordset("Valoracion_Final")
									intPuntuacion = objRecordset("idPuntuacion")
								End if
								objRecordset.MoveNext
							Wend 
						End If
					End If 
				End If 

				'Guardamos la valoración en la tabla del cuestionario
				sql = "UPDATE ECO06_AUTODIAG_VALORACIONES SET Resultado = '" & strNota & "', Puntuacion = " & intNota & ", Puntuacion_Cuestionario = " & intPuntuacion & " WHERE Consulta = "&consulta&" AND Cuestionario = "&cuestionario

				Set objRecordset = Server.CreateObject ("ADODB.Recordset")
				set objRecordset = OBJConnection.Execute(sql)

				'Si el cuestionario no está aprobado, escribimos la valoración en rojo
				strClass = "texto"
				If strNota = "Deficiente" Then
					strClass = "textorojo"
				End if
			%>
				<tr><td class="texto" colspan=2 style="padding-left:15px;">La valoración obtenida en el cuestionario es <span class="<%=strClass%>"><%=strNota%></span>.</td></tr>		
			<%		
					
				
			Else
				'0 respuestas
			End if
	%>
			</table>
			<P align="center"><input type="button" value="REALIZAR OTROS CUESTIONARIOS" class="boton" onclick="location.href='autodiagnostico.asp?consulta=<%=consulta%>'">&nbsp;<input type="button" value="VOLVER A REALIZAR EL CUESTIONARIO" class="boton" onclick="location.href='cuestionario.asp?Cuestionario=<%=cuestionario%>&bloque=<%=bloque%>&consulta=<%=consulta%>'"></P>
	<%
		Else
		
			'Recorrer las preguntas del cuestionario
			If Consulta <> "" Then

				sql = "SELECT ECO06_AUTODIAG_BLOQUES.idBloque, ECO06_AUTODIAG_BLOQUES.Nombre_Bloque, ECO06_AUTODIAG_CUESTIONARIOS.IdCuestionario, ECO06_AUTODIAG_CUESTIONARIOS.NombreCuestionario FROM         ECO06_AUTODIAG_BLOQUES INNER JOIN ECO06_AUTODIAG_CUESTIONARIOS ON ECO06_AUTODIAG_BLOQUES.idBloque = ECO06_AUTODIAG_CUESTIONARIOS.Bloque WHERE IdCuestionario = " & Cuestionario
				Set objRecordset = Server.CreateObject ("ADODB.Recordset")
				set objRecordset = OBJConnection.Execute(sql)
				
				strNombre_Bloque = ""
				strCuestionario = ""
				strImagen = ""
				strIcono = ""
				If Not objRecordset.EOF Then
					If objRecordset("idBloque") = 1 Then
						strIcono = "icon_organiz.gif"
					ElseIf objRecordset("idBloque") = 2 Then
						strIcono = "icon_bidon.gif"
					ElseIf objRecordset("idBloque") = 3 Then
						strIcono = "icon_aspeclabor2.gif"
					End If 
					strImagen = "<IMG SRC='imagenes/" & strIcono & "' align='absmiddle'>"
					strNombre_Bloque = objRecordset("Nombre_Bloque")
					strCuestionario = objRecordset("NombreCuestionario")
				End if

				
				sql = "SELECT ECO06_AUTODIAG_PREGUNTAS.*, ECO06_AUTODIAG_PREGUNTAS_DETALLE.*, ECO06_AUTODIAG_OPCIONES.* FROM ECO06_AUTODIAG_PREGUNTAS INNER JOIN                    ECO06_AUTODIAG_PREGUNTAS_DETALLE ON ECO06_AUTODIAG_PREGUNTAS.IdPregunta = ECO06_AUTODIAG_PREGUNTAS_DETALLE.Pregunta INNER JOIN                     ECO06_AUTODIAG_OPCIONES ON ECO06_AUTODIAG_PREGUNTAS_DETALLE.Tipo_Pregunta = ECO06_AUTODIAG_OPCIONES.idTipo_Pregunta WHERE (ECO06_AUTODIAG_PREGUNTAS.Cuestionario = " & Cuestionario & ") ORDER BY ECO06_AUTODIAG_PREGUNTAS.Orden_Pregunta, ECO06_AUTODIAG_PREGUNTAS_DETALLE.Orden_Detalle, ECO06_AUTODIAG_OPCIONES.idOpcion"
				Set objRecordset = Server.CreateObject ("ADODB.Recordset")
				set objRecordset = OBJConnection.Execute(sql)
	%>

				<form name="formulario" action="cuestionario.asp?Cuestionario=<%=cuestionario%>&bloque=<%=bloque%>&consulta=<%=consulta%>&Respuesta=Si" method="POST">
				<table align="center" width="95%" cellpadding=0 cellspacing=5 border=0 style="background-color: #EFEFEF;" class="tabla2"><tr><td>
				<table style="background: url(imagenes/buscador.gif); background-repeat: no-repeat; background-position: top left; color: #EFEFEF;"><tr><td>
				<table width="95%" cellpadding=0 cellspacing=5 border=0>
				<tr><td class="texto" width="80%"><B><U>Bloque temático: <%=strNombre_Bloque%></U></B></td><td rowspan="2" align="right"><%=strImagen%></td></tr>
				<tr><td class="texto" width="80%"><B><U>Cuestionario: <%=strCuestionario%></U></B></td></tr>
				<tr><td class="texto" colspan=2>&nbsp;</td></tr>
				</table>
				<!-- <table width="95%" cellpadding=0 cellspacing=5 border=0> -->
	<%
				'Buscamos las respuestas anteriores (si ya se hizo el cuestionario) para ponerlas
				sql2 = "SELECT * FROM ECO06_AUTODIAG_RESPUESTAS WHERE Consulta = "&consulta&" AND Cuestionario = "&cuestionario&" ORDER BY Pregunta_detalle, Opcion"

				Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
				set objRecordset2 = OBJConnection.Execute(sql2)
				
				j = 0
				strPregunta = ""
				strDetalle_old = ""
				strCheck = ""
				strValue = ""
				While Not objRecordset.EOF
					if Not objRecordset2.EOF Then
						comprobar_respuesta = "ok"
						intPregunta = objRecordset2("Pregunta_Detalle")
						intOpcion = objRecordset2("Opcion")
						intValor = objRecordset2("Valor")
					Else
						comprobar_respuesta = ""
						strCheck = ""
						strValue = ""
					End If
					MoverRegistro = ""
					
					'Si es visible se muestra siempre
					If CStr(objRecordset("Visible")) = "1" Then
						strDiv = "<table id='capa" & objRecordset("idPregunta_Detalle") & "' width='95%' cellpadding=0 cellspacing=5 border=0 class='visible_tab'>"
					Else
						'Miramos a ver si depende de otra y ocultamos o mostramos
						strDiv = DevuelveVisible(CStr(objRecordset("Condicion_Pregunta")),CStr(objRecordset("Condicion_Respuesta")))
						If strDiv = "hidden" then
							strDiv = "<table id='capa" & objRecordset("idPregunta_Detalle") & "' width='95%' cellpadding=0 cellspacing=5 border=0 class='invisible_tab'>"
						Else
							strDiv = "<table id='capa" & objRecordset("idPregunta_Detalle") & "' width='95%' cellpadding=0 cellspacing=5 border=0 class='visible_tab'>"					
						End if
					End if
					If strPregunta <> objRecordset("Texto_Pregunta") Then
						i = 0
						If j > 0 Then
							response.write "</table>"
						End If
						j = j + 1
						response.write strDiv & "<tr><td class='texto' colspan=2><B>" & objRecordset("Orden_Pregunta") & ". " & sustituyeAyuda(objRecordset("Texto_Pregunta")) & "</B></td></tr>"
					End If
					strDetalle = Trim(objRecordset("Texto_Detalle"))
					If strDetalle_old <> strDetalle And strDetalle <> "" Then
						'If strDetalle_old <> "" Then
							response.write "</table>"
						'End if
						i = i + 1
						response.write strDiv & "<tr><td class='texto' colspan=2 style='padding-left:60px;'>" & i & ". " & sustituyeAyuda(objRecordset("Texto_Detalle")) & "</td></tr>"
					End If
					'Con esta variable escribimos el javascript para msotrar/ocultar celdas de forma dinamica
					strOncheck = ""
					If objRecordset("TextBox") = 0  Then 
						If comprobar_respuesta <> "" Then
							If CStr(objRecordset("idPregunta_Detalle")) = CStr(intPregunta) And CStr(objRecordset("IdOpcion")) = CStr(intValor) Then
								strCheck = "checked"
								MoverRegistro = "ok"
							Else
								strCheck = ""
							End if
						End If
						strOncheck = DevuelveOncheck(CStr(objRecordset("idPregunta_Detalle")),CStr(objRecordset("IdOpcion")))
						response.write "<tr><td class='texto' style='padding-left:30px;' width='30%' align='right'><input type='radio' name='rad_" & objRecordset("idPregunta_Detalle") & "' value='" & objRecordset("IdOpcion") & "' "&strCheck&" "&strOncheck&"></td><td class='texto' align='left'>" & sustituyeAyuda(objRecordset("Texto_Opcion")) & "</td></tr>"
					elseIf objRecordset("TextBox") = 4  Then 	
						If comprobar_respuesta <> "" Then
							If CStr(objRecordset("idPregunta_Detalle")) = CStr(intPregunta) And CStr(objRecordset("IdOpcion")) = CStr(intValor) Then
								strCheck = "checked"
								MoverRegistro = "ok"
							Else
								strCheck = ""
							End if
						End If
						strOncheck = DevuelveOncheck(CStr(objRecordset("idPregunta_Detalle")),CStr(objRecordset("IdOpcion")))
						response.write "<tr><td class='texto' style='padding-left:30px;' width='30%' align='right'>&nbsp;</td><td class='texto' align='left' style='padding-left:30px;'><input type='radio' name='rad_" & objRecordset("idPregunta_Detalle") & "' value='" & objRecordset("IdOpcion") & "' "&strCheck&" "&strOncheck&">&nbsp;" & sustituyeAyuda(objRecordset("Texto_Opcion")) & "</td></tr>"
					ElseIf objRecordset("TextBox") = 2 Then
						If comprobar_respuesta <> "" Then
							If CStr(objRecordset("idPregunta_Detalle")) = CStr(intPregunta) And CStr(objRecordset("IdOpcion")) = CStr(intValor) Then
								strCheck = "checked"
								MoverRegistro = "ok"
							Else
								strCheck = ""
							End if
						End if
						response.write "<tr><td class='texto' style='padding-left:30px;' width='30%' align='right'><input type='checkbox' name='chk_" & objRecordset("idPregunta_Detalle") & "_" & objRecordset("IdOpcion") & "' "&strCheck&"></td><td class='texto' align='left'>" & sustituyeAyuda(objRecordset("Texto_Opcion")) & "</td></tr>"
					ElseIf objRecordset("TextBox") = 1 Then
						If comprobar_respuesta <> "" Then
							If CStr(objRecordset("idPregunta_Detalle")) = CStr(intPregunta) And CStr(objRecordset("IdOpcion")) = CStr(intOpcion) Then
								strValue = CStr(intValor)
								MoverRegistro = "ok"
							Else
								strValue = ""
							End if
						End if
						response.write "<tr><td class='texto' style='padding-left:30px;' width='30%' align='right'>" & sustituyeAyuda(objRecordset("Texto_Opcion")) & "</td><td class='texto' align='left'><input type='textbox' name='txt_" & objRecordset("idPregunta_Detalle") & "_" & objRecordset("IdOpcion") & "' value='"&strValue&"' size='20' maxlenght='20'></td></tr>"
					ElseIf objRecordset("TextBox") = 3 Then
						response.write "<tr><td class='texto' style='padding-left:30px;' width='30%' align='right'><b>-</b></td><td class='texto' align='left'><b>" & sustituyeAyuda(objRecordset("Texto_Opcion")) & "</b></td></tr>"
					End If 
					strDetalle_old = strDetalle
					strPregunta = objRecordset("Texto_Pregunta")
					If MoverRegistro <> "" Then
						objRecordset2.MoveNext
					End if
					objRecordset.MoveNext	
				Wend 
				
				response.write "</table>"	
					%>
					
			<!-- 		</td></tr>
				</table>  -->
				</td></tr></table>
				</td></tr>
				<tr><td align=center>&nbsp;</td></tr>
				<tr><td align=center><input type="button" value="GUARDAR" class="boton" onclick="comprueba_pagina()"></td></tr>
				<tr><td align=center>&nbsp;</td></tr>
				<tr><td align=left>(*) Es obligatorio responder todas las preguntas.</td></tr>
				</table>
			<!-- <CENTER><input type="submit" value="BUSCAR" class="boton" onclick="document.buscador.submit()"> </CENTER> -->

			</form>
			
		<% End If  %>
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