<!--#include file="web_inicio.asp"-->
<%

idcarpeta_bv = EliminaInyeccionSQL(request("id"))

sql = "SELECT nombre,numeracion FROM ENL_TEMAS where id="&idcarpeta_bv
set objRecordset = Server.CreateObject ("ADODB.Recordset")
objRecordset.Open sql,objConnection,adOpenKeyset
tema = objRecordset("nombre")

	If idcarpeta_bv <> "" then
		sql = "SELECT ENL_ENLACES.* FROM ENL_ENLACES LEFT JOIN ENL_CLASIFICACION ON ENL_ENLACES.id=ENL_CLASIFICACION.enlace WHERE ENL_CLASIFICACION.tema="&idcarpeta_bv&" AND (ENL_ENLACES.Permiso=52 or ENL_ENLACES.permiso=0) ORDER BY ENL_CLASIFICACION.orden"
		set obj = Server.CreateObject ("ADODB.Recordset")
		obj.Open sql,objConnection,adOpenKeyset		
		cuantos_enlaces = obj.RecordCount
			do while not obj.eof  
			  'Preparar minificha y ventana con mas detalles
			  strMinificha = ""
			  strTextoFicha = ""
			  strTitulo = ""
			  strSubtitulo = ""
			  strDescripcion = ""
			  strAfiliacion = ""
			  If Trim(obj("titulo")) <> "" And Not IsNull(obj("titulo")) Then
					strTitulo = obj("titulo")
					strTitulo = Limpia(strTitulo)
			  End If 
			  If Trim(obj("subtitulo")) <> "" And Not IsNull(obj("subtitulo")) Then
					strSubtitulo = obj("subtitulo")
					strSubtitulo = Limpia(strSubtitulo)
			  End If 
			  If Trim(obj("descripcion")) <> "" And Not IsNull(obj("descripcion")) Then
					strDescripcion = obj("descripcion")
					strDescripcion = Limpia(strDescripcion)
			  End If
			  strSubtituloDescripcion = strSubtitulo
			  If strSubtituloDescripcion="" then strSubtituloDescripcion = strDescripcion

			  If Trim(obj("afiliacion")) <> "" And Not IsNull(obj("afiliacion")) Then
					strAfiliacion = obj("afiliacion")
					strAfiliacion = Limpia(strAfiliacion)
			  End If		
		
			  idfruto = obj("id")
		  	  afiliacion = DamePublicacion(idfruto)
			  
			  if afiliacion = "ISTAS" then
					cuantas_publicaciones = cuantas_publicaciones+1
					if not(isnull(obj("imagen"))) and obj("imagen")<>"" then 
						im_pub = "http://www.istas.net/webistas/imagenes/"&obj("imagen")
					else
						im_pub = "img/noimagenpublic.gif"
					end If
					im_CCOO = ""
					strPub = "ISTAS"
					strMinificha = "Abrir Publicación ISTAS"
					strResultado = strResultado & "<div class=frutos_publicaciones"&fuente&"><table width=100% cellpadding=0 cellspacing=0><tr><td valign='top' align='right' width='55px'>"
					strResultado = strResultado & "<a onclick=window.open('wi_ficha_enlace.asp?tipo_fruto="&tipo_fruto&"&idfruto="&idfruto&"&idpagina="&idpagina&"','def','width=400,height=400,scrollbars=yes,resizable=yes') onkeypress=window.open('wi_ficha_enlace.asp?tipo_fruto="&tipo_fruto&"&idfruto="&idfruto&"&idpagina="&idpagina&"','def','width=400,height=400,scrollbars=yes,resizable=yes') style='cursor:pointer'>"
					if cstr(request("lugar"))<>"2" then
						strResultado = strResultado & "<img src="&im_pub&" WIDTH='46px' HEIGHT='64px' border=0 align=left alt='Detalles de la Publicación " & strPub & ": " & strTitulo & "'></a></td><td valign='middle' style='padding:4px'>" & im_CCOO & "<a class='indexlink"&fuente&"' href='abreenlace.asp?idenlace=" & idfruto & "' title='"&strMinificha&"' target='_blank'><b>" & strTitulo & "</b>"
						if strSubtituloDescripcion<>"" then strResultado = strResultado &  "<br>"&strSubtituloDescripcion
						strResultado = strResultado &  "</a></td></tr></table></div>"
					else
						strResultado = strResultado & "<img src="&im_pub&" WIDTH='46px' HEIGHT='64px' border=0 align=left alt='Detalles de la Publicación " & strPub & ": " & strTitulo & "'></a></td><td valign='middle' style='padding:4px'>" & im_CCOO & "<a class='indexlink"&fuente&"' href='abreenlace.asp?idenlace=" & idfruto & "' title='"&strMinificha&"' target='_blank'>" & strTitulo
						strResultado = strResultado &  "</a></td></tr></table></div>"
					end if
			  else
					cuantas_publicaciones = cuantas_publicaciones+1
					If afiliacion = "CCOO" Then
						im_CCOO = "<IMG SRC='ccoo.gif' WIDTH='23' HEIGHT='8' BORDER='0' ALT='Publicación de CCOO'>&nbsp;"
						strPub = "CCOO"
						strMinificha = "Abrir Publicación CC.OO."
						strResultado = strResultado & "<table width=100% cellpadding=0 cellspacing=0><tr><td valign='top' style='padding-top:6px' align='right' width='25px'>"
						strResultado = strResultado & "<a onclick=window.open('wi_ficha_enlace.asp?tipo_fruto="&tipo_fruto&"&idfruto="&idfruto&"&idpagina="&idpagina&"','def','width=400,height=400,scrollbars=yes,resizable=yes') onkeypress=window.open('wi_ficha_enlace.asp?tipo_fruto="&tipo_fruto&"&idfruto="&idfruto&"&idpagina="&idpagina&"','def','width=400,height=400,scrollbars=yes,resizable=yes') style='cursor:pointer'>"
						if cstr(request("lugar"))<>"2" then
							strResultado = strResultado & "<img src=enlace.gif WIDTH='16px' HEIGHT='12px' border=0 alt='Detalles de la Publicación " & strPub & "'></a></td><td valign='middle' style='padding:4px'>" & im_CCOO & "<a class='indexlink"&fuente&"' href='abreenlace.asp?idenlace=" & idfruto & "' title='"&strMinificha&"' target='_blank'><b>" & strTitulo & "</b>"
							if strSubtituloDescripcion<>"" then strResultado = strResultado &  "<br>"&strSubtituloDescripcion
							strResultado = strResultado &  "</a></td></tr></table>"
						else
							strResultado = strResultado & "<img src=enlace.gif WIDTH='16px' HEIGHT='12px' border=0 alt='Detalles de la Publicación " & strPub & "'></a></td><td valign='middle' style='padding:4px'>" & im_CCOO & "<a class='indexlink"&fuente&"' href='abreenlace.asp?idenlace=" & idfruto & "' title='"&strMinificha&"' target='_blank'>" & strTitulo
							strResultado = strResultado &  "</a></td></tr></table>"
						end if
					else
						if cstr(obj("doc_web"))="1" then strMinificha = "Visitar la página"
						if cstr(obj("doc_fichero"))="1" then strMinificha = "Abrir el fichero"
						strResultado = strResultado & "<table width=100% cellpadding=0 cellspacing=0><tr><td valign='top' style='padding-top:6px' align='right' width='25px'>"
						strResultado = strResultado & "<a onclick=window.open('wi_ficha_enlace.asp?tipo_fruto="&tipo_fruto&"&idfruto="&idfruto&"&idpagina="&idpagina&"','def','width=400,height=400,scrollbars=yes,resizable=yes') onkeypress=window.open('wi_ficha_enlace.asp?tipo_fruto="&tipo_fruto&"&idfruto="&idfruto&"&idpagina="&idpagina&"','def','width=400,height=400,scrollbars=yes,resizable=yes') style='cursor:pointer'>"
						strResultado = strResultado & "<img src=enlace.gif WIDTH='16' HEIGHT='12' border=0 alt='Detalles del enlace'></a></td><td valign='middle' style='padding:4px' class=cuerpo><a class='indexlink"&fuente&"' href='abreenlace.asp?idenlace=" & idfruto & "' title='"&strMinificha&"' target='_blank'><b>" & strAfiliacion & "</b>&nbsp;&gt;&nbsp;" & strTitulo & "</a><br>" & strSubtitulo & "</td></tr></table>"
					end if
			  end if

	  obj.movenext
	  loop
	  if cuantos_enlaces=0 and cstr(lugar) = "0" then strResultado="<font class=cuerpo>(no hay enlaces, elige un subtema)</font>"
	End If 

Function DamePublicacion(id_Enlace)
	strResultadoPublic = "no"
	
	sql = "SELECT ENL_TEMAS.id,nombre, numeracion, enlace FROM ENL_TEMAS INNER JOIN ENL_CLASIFICACION ON ENL_TEMAS.id = ENL_CLASIFICACION.tema WHERE ((numeracion LIKE '" & str_NumCCOO & "%') OR (numeracion LIKE '" & str_NumISTAS & "%')) and enlace = " & id_Enlace

	set objRecordset33 = Server.CreateObject ("ADODB.Recordset")
	objRecordset33.Open sql,objConnection,adOpenKeyset

	If Not objRecordset33.EOF Then
		If str_NumCCOO <> "" then
			If Left(objRecordset33("numeracion"),Len(str_NumCCOO)) = str_NumCCOO Then
				strResultadoPublic = "CCOO"
			End If
		End If 
		If str_NumISTAS <> "" then
			If Left(objRecordset33("numeracion"),Len(str_NumISTAS)) = str_NumISTAS Then
				strResultadoPublic = "ISTAS"
			End If
		End If 
	End If
	
	DamePublicacion = strResultadoPublic

End Function 

Function DameOtrasRamas(int_Fruto,int_TipoFruto)

	'Buscamos además en qué otros temas está clasificado el fruto
	sql = "SELECT WI_ARBOL.Numeracion, WI_ARBOL.Titulo FROM WI_FRUTOS INNER JOIN WI_ARBOL ON WI_FRUTOS.idRama = WI_ARBOL.IdRama WHERE (WI_FRUTOS.idFruto = "&int_Fruto&") AND (WI_FRUTOS.Tipo_Fruto = "&int_TipoFruto&") AND (WI_ARBOL.IdRama <> "&idpagina&") GROUP BY WI_ARBOL.Numeracion, WI_ARBOL.Titulo, WI_ARBOL.IdRama ORDER BY WI_ARBOL.Numeracion"
	set objRecordsetRama = Server.CreateObject ("ADODB.Recordset")
	objRecordsetRama.Open sql,objConnection,adOpenKeyset
	str_texto_Minificha = ""
	str_texto_Minificha = str_texto_Minificha & "Otros temas en los que está clasificado este enlace:"
	While Not objRecordsetRama.EOF 
		strNumeracion = objRecordsetRama("Numeracion")
		strAux = ""
		For intCaracter = 2 To Len(strNumeracion)
			strAux = strAux & CStr(Asc(Mid(strNumeracion,intCaracter,1)) - 64) & "."
		Next 
		strNumeracion = strAux
		strTitulo = objRecordsetRama("Titulo")
		str_texto_Minificha = str_texto_Minificha & vbcrlf & strNumeracion & " " & strTitulo
		objRecordsetRama.MoveNext
	wend
	
	DameOtrasRamas = str_texto_Minificha

End Function

%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: enlaces biblioteca virtual</title>
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

<body topmargin="0" leftmargin="0">
<table align="center" width="772" border="0" cellpadding="0" cellspacing="0" style="background-color: #00AC5A">
<tr>
	<td width="160"><img src="enlaces.gif" border="0" alt="Logo ECOinformas"></td>
	<td width="612" class="textoblanco" align="right">Enlaces suministrado por la biblioteca virtual de <a href="http://www.ecoinformas.com/" style="text-decoration:none" target="_top">ECOinformas</a>&nbsp;<br><%=url%>&nbsp;</td>
</tr>
</table>
<%=strResultado%>
</body>