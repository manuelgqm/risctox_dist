<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

	
	
	'----- Si es restringida y no estás identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	id_ecogente = session("id_ecogente")
	'---- ATENCIÓN: ponerlo cuando publiquemos en abierto
	
	numeracion = "AICCA"
		
	FUNCTION formato(x,lon)
		if isnull(x) then
			formato = ""
		else
			'x = replace(x,chr(10),"<br>")
			x = ucase(x)
			x = replace(x,"ACUTE;","acute;")
			if len(x)>(lon-3) then x = mid(x,1,lon-3)&"..."
			formato = x
		end if
	END FUNCTION


FUNCTION unQuote(s)
  pos = Instr(s, "'")
  While pos > 0 
    s = Mid(s,1,pos) & "'" & Mid(s,pos+1)
    pos = InStr(pos+2, s, "'")
  Wend
  pos = Instr(s, """")
  While pos > 0 
    s = Mid(s,1,pos-1) & "''" & Mid(s,pos+1)
    pos = InStr(pos+2, s, """")
  Wend
  unQuote = Trim(s)
END FUNCTION
	
%>

<html>
<head>
<title>ECOinformas: Base de datos de sustancias tóxicas y peligrosas RISCTOX</title>
<link rel="stylesheet" type="text/css" href="estructura.css"  />
<SCRIPT LANGUAGE="JavaScript">
<!--
function imprimir() 
{

		if  (confirm('¿Imprimir este documento?')) { print(); close();}
}

// -->
</SCRIPT>
</head>
<body onload="imprimir()">
<table style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; color: #000000;"><tr><td>
<!--identificación-->
				<table width="100%" cellpadding=5><tr><td><a name="identificacion"></a><img src="imagenes/risctox01.gif" alt="identificación de la sustancia" width="255" height="32" /></td>
		                <td align="right"></td></tr></table>

<%							
				CAS_actual = trim(request("CAS"))
				sql = "SELECT * FROM RQ_SUSTANCIAS WHERE CAS='"&CAS_actual&"'"
				set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset.Open sql,objConnection,adOpenKeyset
		   	   	if not objRecordset.eof then %>
<%	'-- 1: DATOS SUSTANCIA											%>
				
		   	   	<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		   	   	<tr><td class="celdaabajo" colspan="2" align="center">SUSTANCIA</td></tr>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Nombre:</td><td class="texto" valign="middle"><b><%=formato(objRecordset("nombre"),300)%></b></td></tr>
		   	   	<% if not isnull(objRecordset("sinonimos")) and objRecordset("sinonimos")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Sinónimos:</td><td class="texto" valign="middle"><b><%=formato(objRecordset("sinonimos"),300)%></b></td></tr>
		   	   	<% end if %>
		   	   	<% if not isnull(objRecordset("cas")) and objRecordset("cas")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">CAS:</td><td class="texto" valign="middle"><%=formato(objRecordset("cas"),300)%></td></tr>
		   	   	<% end if %>
		   	   	<% if not isnull(objRecordset("cee")) and objRecordset("cee")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Nº C.E./EINECS:</td><td class="texto" valign="middle"><%=formato(objRecordset("cee"),300)%></td></tr>
		   	   	<% end if %>
		   	   	<% if not isnull(objRecordset("rd")) and objRecordset("rd")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top" nowrap>R.D. 363/1995:</td><td class="texto" valign="middle"><%=formato(objRecordset("rd"),300)%></td></tr>
				<% else %>
				<tr><td class="texto" valign="middle" colspan="2" align="center">Sustancia no incluida en el Anexo I del RD 363/1995.<br>Es responsabilidad del fabricante de la sustancia o preparado asignarle las Frases R y S</td></tr>
				<% end if %>
				</table>
				
				<div style="height:3pt"></div>
<%	'-- 2: CLASIFICACIÓN 										  %>

				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">CLASIFICACIÓN</td></tr>
		   	   	<% if not isnull(objRecordset("simbolo")) and objRecordset("simbolo")<>"" then %>
				<tr><td class="subtitulo3" align="right" valign="top">Símbolos:</td><td class="texto" valign="middle"><%=objRecordset("simbolo")%></td></tr>
				<tr><td class="texto" valign="middle" colspan="2" align="center">
				<% simbolos = replace(objRecordset("simbolo"),";",",")
				   'dim simbolo(30)
				   simbolo = split(simbolos,",")
				   for i=0 to Ubound(simbolo)
				   	if ucase(mid(trim(simbolo(i)),1,1))="T" then response.write "<img src='http://www.istas.net/recursos/VEC/ISTAS_00324.wmf' width=100 alt='Tóxicos'>&nbsp;"
				   	if ucase(mid(trim(simbolo(i)),1,1))="C" then response.write "<img src='http://www.istas.net/recursos/VEC/ISTAS_00318.wmf' width=100 alt='Corrosivos'>&nbsp;"
				   	if ucase(mid(trim(simbolo(i)),1,1))="N" then response.write "<img src='http://www.istas.net/recursos/VEC/ISTAS_00323.wmf' width=100 alt='Peligrosos para el medio ambiente'>&nbsp;"
				   	if ucase(mid(trim(simbolo(i)),1,1))="F" then response.write "<img src='http://www.istas.net/recursos/VEC/ISTAS_00320.wmf' width=100 alt='Inflamables'>&nbsp;"
				   	if ucase(mid(trim(simbolo(i)),1,1))="X" then response.write "<img src='http://www.istas.net/recursos/VEC/ISTAS_00321.wmf' width=100 alt='Nocivos'>&nbsp;"
				   	if ucase(mid(trim(simbolo(i)),1,1))="O" then response.write "<img src='http://www.istas.net/recursos/VEC/ISTAS_00314.wmf' width=100 alt='Comburentes'>&nbsp;"
				   	if ucase(mid(trim(simbolo(i)),1,1))="E" then response.write "<img src='http://www.istas.net/recursos/VEC/ISTAS_00319.wmf' width=100 alt='Explosivos'>&nbsp;"
				   next %>

				</td></tr>
				<% end if %>
				<% for i=1 to 11 
					campo = "clasific"&cstr(i)
					if not isnull(objRecordset(campo)) and objRecordset(campo)<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Frases R (clasificación <%=i%>):</td><td class="texto" valign="middle"><%=objRecordset(campo)%>
		   	   	</td></tr>
		   	   	<% 	end if %>
		   	   	<% next %>
		   	   	<% if not isnull(objRecordset("frases_s")) and objRecordset("frases_s")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Frases S:</td><td class="texto" valign="middle"><%=objRecordset("frases_s")%>&nbsp;
		   	   	</td></tr>
				<% end if %>
				</table>
				
				<div style="height:3pt"></div>

<%	'-- 3: ETIQUETADO
				   eticonc= ""
				   for i=1 to 11
					if not isnull(objrecordset("eticonc"&cstr(i))) then eticonc = eticonc & objrecordset("eticonc"&cstr(i))
				   next
				   if eticonc<>"" then %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">ETIQUETADO</td></tr>
				<% for i=1 to 11 
					campo = "eticonc"&cstr(i)
					campo2 = "conc"&cstr(i)
					if not isnull(objRecordset(campo)) and objRecordset(campo)<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top" nowrap>Concentración <%=i%>:</td>
		   	   	    <td class="texto" valign="middle"><%=objRecordset(campo2)%></td>
		   	   	    <td class="subtitulo3" align="right" valign="top" nowrap>Etiqueta <%=i%>:</td>
		   	   	    <td class="texto" valign="middle"><%=objRecordset(campo)%>&nbsp;</td>
		   	   	</tr>
		   	   	<% 	end if %>
		   	   	<% next %>
				</table>

				<div style="height:3pt"></div>
				<% end if %>
<% else %>
<%	'-- 1 danés: DATOS SUSTANCIA											%>
<%				sql2 = "SELECT * FROM RQ_SUSTANCIAS_DANESAS WHERE CAS='"&CAS_actual&"'"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then %>
				
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		   	   	<tr><td class="celdaabajo" colspan="2" align="center">SUSTANCIA</td></tr>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Nombre:</td><td class="texto" valign="middle"><b><%=formato(request("nombre"),300)%></b></td></tr>
		   	   	<% if not isnull(objRecordset2("cas")) and objRecordset2("cas")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">CAS:</td><td class="texto" valign="middle"><%=formato(objRecordset2("cas"),300)%></td></tr>
		   	   	<% end if %>
		   	   	<% if not isnull(objRecordset2("einecs")) and objRecordset2("einecs")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Nº C.E./EINECS</td><td class="texto" valign="middle"><%=formato(objRecordset2("einecs"),300)%></td></tr>
		   	   	<% end if %>
		   	   	<tr><td class="texto" valign="middle" colspan="2" align="center">Sustancia no incluida en el Anexo I del RD 363/1995.<br>Es responsabilidad del fabricante de la sustancia o preparado asignarle las Frases R y S</td></tr>
				</table>

				<div style="height:3pt"></div>
<%	'-- 2 danés: CLASIFICACIÓN 										  %>

				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">CLASIFICACIÓN</td></tr>
				<% if not isnull(objRecordset2("frases_r")) and objRecordset2("frases_r")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Frases R (recomendadas por la Agencia de Medio Ambiente de Dinamarca):</td><td class="texto" valign="middle"><%=objRecordset2("frases_r")%>&nbsp;
		   	   	</td></tr>
		   	   	<% end if %>
				</table>

				<div style="height:3pt"></div>

				<% else %>

<%	'-- 1 no está en RD ni en la lista danesa: DATOS SUSTANCIA											%>
		   	   	
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		   	   	<tr><td class="celdaabajo" colspan="2" align="center">SUSTANCIA</td></tr>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Nombre:</td><td class="texto" valign="middle"><b><%=formato(request("nombre"),300)%></b></td></tr>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">CAS:</td><td class="texto" valign="middle"><%=formato(request("cas"),300)%></td></tr>
				<tr><td class="texto" valign="middle" colspan="2" align="center">Sustancia no incluida en el Anexo I del RD 363/1995.<br>Es responsabilidad del fabricante de la sustancia o preparado asignarle las Frases R y S</td></tr>
				</table>
				<div style="height:3pt"></div>

				<% end if %>
<% end if %>
				<div style="height:3pt"></div>
<!--fin de identificación-->

<% if request("ficha")<>"identificacion" then %>

<!--riesgos para la salud-->				
		                <table width="100%" cellpadding=5><tr><td><a name="riesgossalud"></a><img src="imagenes/risctox02.gif" alt="riesgos específicos para la salud"></td>
		                <td align="right"></td></tr></table>

<% 	'-- 4: Cancerígenas y mutágenas (según RD 363/1995)
				sql2 = "SELECT C,M,notas FROM RQ_SUSTANCIAS_CYM WHERE cas='"&CAS_actual&"' GROUP BY C,M,notas"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then 
		   	   		texto4="si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">CANCERÍGENA Y MUTÁGENA (según RD 363/1995)</td></tr>
		   	   	<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("C")) and objRecordset2("C")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Cancerígeno:</td><td class="texto" valign="middle">
		   	   		<% if objRecordset2("c")="C1" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=1','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>C1</a>"
					   end if
					   if objRecordset2("c")="C2" then
					   	response.write "<a onclick=window.open('ver_definicion.asp?id=2','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>C2</a>"
					   end if %>
		   	   	</td></tr>
				<% end if %>
		   	   	<% if not isnull(objRecordset2("M")) and objRecordset2("M")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Mutágeno:</td><td class="texto" valign="middle">
		   	   		<% if objRecordset2("M")="M2" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=3','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>M2</a>"
					   end if %>
		   	   	</td></tr>
				<% end if %>
		   	   	<% if not isnull(objRecordset2("notas")) and objRecordset2("notas")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Notas:</td><td class="texto" valign="middle">
		   	   		<% if ucase(objRecordset2("notas"))="TR1" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=4','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>TR1</a>"
					   end if
   		   	   		   if ucase(objRecordset2("notas"))="TR2" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=5','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>TR2</a>"
					   end if 
					   if ucase(objRecordset2("notas"))="Q" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=6','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>Q</a>"
					   end if 
					   if ucase(objRecordset2("notas"))="SEN" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=14','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>SEN</a>"
					   end if 
					   if objRecordset2("notas")="véase Tabla 3" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=8','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>VER TABLA</a>"
					   end if %>
		   	   	</td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>

				<div style="height:3pt"></div>
				<% end if %>
				
<% 	'-- 5: Cancerígenas y mutágenas (según IARC)
				sql2 = "SELECT grupo,volumen FROM RQ_SUSTANCIAS_CYM2 WHERE cas='"&CAS_actual&"' GROUP BY grupo,volumen"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto5 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">CANCERÍGENA Y MUTÁGENA (según IARC)</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("grupo")) and objRecordset2("grupo")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Grupo:</td><td class="texto" valign="middle">
		   	   		<% if ucase(objRecordset2("grupo"))="GRUPO 1" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=9','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>Grupo 1</a>"
					   end if
					   if ucase(objRecordset2("grupo"))="GRUPO 2A" then
					   	response.write "<a onclick=window.open('ver_definicion.asp?id=10','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>Grupo 2A</a>"
					   end if 
					   if ucase(objRecordset2("grupo"))="GRUPO 2B" then
					   	response.write "<a onclick=window.open('ver_definicion.asp?id=11','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>Grupo 2B</a>"
					   end if 
					   if ucase(objRecordset2("grupo"))="GRUPO 3" then
					   	response.write "<a onclick=window.open('ver_definicion.asp?id=12','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>Grupo 3</a>"
					   end if 
					   if ucase(objRecordset2("grupo"))="GRUPO 4" then
					   	response.write "<a onclick=window.open('ver_definicion.asp?id=13','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>Grupo 4</a>"
					   end if %>
		   	   	</td></tr>
				<% end if %>
				
		   	   	<% if not isnull(objRecordset2("volumen")) and objRecordset2("volumen")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Volumen:</td><td class="texto" valign="middle"><%=objRecordset2("volumen")%></td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>

				<div style="height:3pt"></div>
				<% end if %>
				

<% 	'-- 6: Cancerígenas y mutágenas (según otras fuentes)
				sql2 = "SELECT fuente FROM RQ_SUSTANCIAS_CYM3 WHERE cas='"&CAS_actual&"'"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then 
		   	   		texto6 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">CANCERÍGENA Y MUTÁGENA (según otras fuentes)</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("fuente")) and objRecordset2("fuente")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Fuentes:</td><td class="texto" valign="middle">
		   	   		<% fuentes = split(objRecordset2("fuente"),",")
					   for i=0 to Ubound(fuentes)
					   
					   if trim(ucase(fuentes(i)))="O" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=15','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>O</a>"
					   end if
					   if trim(ucase(fuentes(i)))="G-A1" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=16','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>G-A1</a>"
					   end if
					   if trim(ucase(fuentes(i)))="G-A2" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=17','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>G-A2</a>"
					   end if
					   if trim(ucase(fuentes(i)))="G-A3" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=18','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>G-A3</a>"
					   end if
					   if trim(ucase(fuentes(i)))="G-A4" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=19','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>G-A4</a>"
					   end if
					   if trim(ucase(fuentes(i)))="G-A5" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=20','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>G-A5</a>"
					   end if
					   if trim(ucase(fuentes(i)))="N-1" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=21','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>N-1</a>"
					   end if
					   if trim(ucase(fuentes(i)))="N-2" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=22','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>N-2</a>"
					   end if
					   if trim(ucase(fuentes(i)))="CP65" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=23','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>CP65</a>"
					   end if
					   response.write "&nbsp;&nbsp;"
					   next %>
		   	   	</td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>

				<div style="height:3pt"></div>
				<% end if %>

<% if id_ecogente=179 then %>
<% 	'-- 0: Tóxico para la reproducción
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS WHERE ((RD<>'' OR isnull(RD,'nulo')<>'nulo') AND (1=0"
				for i=1 to 11
					sql2 = sql2 &" OR clasific"&i&" LIKE '%60%' OR clasific"&i&" LIKE '%61%' OR clasific"&i&" LIKE '%62%' OR clasific"&i&" LIKE '%63%'" 
				next
				sql2 = sql2 & ") AND (cas='"&CAS_actual&"'))"
				'sql2 = "SELECT fuente FROM RQ_SUSTANCIAS_CYM3 WHERE cas='"&CAS_actual&"'"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto0 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<font class="titulo3">TÓXICO PARA LA REPRODUCCIÓN</font>
					</td><td width="20%" align="right">
					&nbsp;
					</td></tr></table>
				</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>
<% end if %>				



<% 	'-- 7: Disruptores endocrinos
				sql2 = "SELECT fuente FROM RQ_SUSTANCIAS_DIS WHERE cas='"&CAS_actual&"' GROUP BY fuente"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto7 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">DISRUPTOR ENDOCRINO</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("fuente")) and objRecordset2("fuente")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Fuente:</td><td class="texto" valign="middle">
		   	   		<% if trim(ucase(objRecordset2("fuente")))="NS" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=24','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>NS</a>"
					   end if
					   if trim(ucase(objRecordset2("fuente")))="UE1" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=25','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>UE1</a>"
					   end if
					   if trim(ucase(objRecordset2("fuente")))="UE2" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=26','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>UE2</a>"
					   end if %>
		   	   	</td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% if id_ecogente=179 then %>
<% 	'-- 8: Neurotóxicos
				sql2 = "SELECT efecto,nivel,fuente FROM RQ_SUSTANCIAS_NEU WHERE cas='"&CAS_actual&"'"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto8 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">NEUROTÓXICO</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("fuente")) and objRecordset2("fuente")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Efecto:</td><td class="texto" valign="middle">
					<% 
					   if trim((objRecordset2("efecto")))="SNC" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=75','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>SNC</a>"
					   else
					     if trim((objRecordset2("efecto")))="SNP" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=76','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>SNP</a>"
					     else
					   	response.write formato(objRecordset2("efecto"),25)
					     end if
					   end if %>
		   	   	</td></tr>
				<% end if %>
		   	   	<% if not isnull(objRecordset2("nivel")) and objRecordset2("nivel")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Nivel:</td><td class="texto" valign="middle">
					<% 
					   if trim((objRecordset2("nivel")))="1" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=77','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>1</a>"
					   end if
					   if trim((objRecordset2("nivel")))="2" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=78','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>2</a>"
					   end if
					   if trim((objRecordset2("nivel")))="3" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=79','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>3</a>"
					   end if
					   if trim((objRecordset2("nivel")))="4" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=80','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>4</a>"
					   end if %>
		   	   	</td></tr>
				<% end if %>
		   	   	<% if not isnull(objRecordset2("fuente")) and objRecordset2("fuente")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Fuentes:</td><td class="texto" valign="middle">
					<% fuentes = split(objRecordset2("fuente"),",")
					   for i=0 to Ubound(fuentes)
						response.write "<a onclick=window.open('ver_definicion.asp?id="&clng(fuentes(i))+50&"','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>"&fuentes(i)&"</a>&nbsp;&nbsp;"
					   next %>				
		   	   	</td></tr>
				<% end if %>				
				<% objRecordset2.movenext
				   loop %>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 0b: Sensibilizante
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS WHERE ((RD<>'' OR isnull(RD,'nulo')<>'nulo') AND (1=0"
				for i=1 to 11
					sql2 = sql2 &" OR clasific"&i&" LIKE '%42%' OR clasific"&i&" LIKE '%43%' " 
				next
				sql2 = sql2 & ") AND (cas='"&CAS_actual&"'))"
				'sql2 = "SELECT fuente FROM RQ_SUSTANCIAS_CYM3 WHERE cas='"&CAS_actual&"'"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto0b = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">SENSIBILIZANTE</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>
<% end if %>
<!--fin de riesgos para la salud-->

<!--riesgos para el medioambiente-->
				<div style="height:3pt"></div>
		                <table width="100%" cellpadding=5><tr><td><a name="riesgosma"></a><img src="imagenes/risctox03.gif" alt="riesgos específicos para el medioambiente"></td>
		                <td align="right"></td></tr></table>

<% 	'-- 9: Persistencia y bioacumulación
				sql2 = "SELECT enlace,url FROM RQ_SUSTANCIAS_PYB WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto9 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">TÓXICAS, PERSISTENTES Y BIOACUMULATIVAS</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("url")) and objRecordset2("url")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Más información (en inglés):</td><td class="texto" valign="middle">
					 <a href="<%=lcase(objRecordset2("url"))%>" target="_blank"><%=mid(lcase(objRecordset2("enlace")),1,100)%></a>&nbsp;
		   	   	</td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>

				<div style="height:3pt"></div>
				<% end if %>

<% if id_ecogente=179 then %>
<% 	'-- 10: Toxicidad acuática (según directiva de aguas)
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS_TAC WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto10 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">TOXICIDAD ACUÁTICA (según directiva de aguas)</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 11: Toxicidad acuática (Peligrosas agua Alemania)
				sql2 = "SELECT campo5 FROM RQ_SUSTANCIAS_TAC2 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto11 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">TOXICIDAD ACUÁTICA (Peligrosas agua Alemania)</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("campo5")) and objRecordset2("campo5")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Clasificación:</td><td class="texto" valign="middle">
					<%
					if objRecordset2("campo5")="nwg" then response.write "no peligrosa para aguas"
					if objRecordset2("campo5")="1" then response.write "baja peligrosidad para aguas"
					if objRecordset2("campo5")="2" then response.write "peligrosa para aguas"
					if objRecordset2("campo5")="3" then response.write "elevada peligrosidad para aguas"
					%>
		   	   	</td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>

				<div style="height:3pt"></div>
				<% end if %>
<% end if %>
<% 	'-- 12: Daño a la atmósfera (Capa de Ozono)
				sql2 = "SELECT nombre2 FROM RQ_SUSTANCIAS_DAT WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto12 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">DAÑO A LA ATMÓSFERA (Capa de Ozono)</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("nombre2")) and objRecordset2("nombre2")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Otro nombre:</td>
				<td class="texto" valign="middle"><%=formato(objRecordset2("nombre2"),100)%></td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 13: Daño a la atmósfera (cambio climático)
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS_DAT2 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto13 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">DAÑO A LA ATMÓSFERA (cambio climático)</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 14: Daño a la atmósfera (calidad del aire)
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS_DAT3 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto14 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">DAÑO A LA ATMÓSFERA (calidad del aire)</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

				</div>
<!--fin de riesgos para el medioambiente-->

<% if id_ecogente=179 then %>
<!--normativa salud laboral-->
				<div style="height:3pt"></div>

		                <table width="100%" cellpadding=5><tr><td><a name="normativasalud"></a><img src="imagenes/risctox04.gif" alt="normativa salud laboral"></td>
		                <td align="right"></td></tr></table>


<% 	'-- 15: Valores límite (VLA)
				sql2 = "SELECT vlaed1,vlaed2,vlaec1,vlaec2,notas FROM RQ_SUSTANCIAS_VL1 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto15 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">VALORES LÍMITE AMBIENTALES</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("vlaed1")) and objRecordset2("vlaed1")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">VLA-ED:</td>
				<td class="campo" valign="middle"><%=formato(objRecordset2("vlaed1"),100)%>&nbsp;ppm</td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("vlaed2")) and objRecordset2("vlaed2")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">VLA-ED:</td>
				<td class="campo" valign="middle"><%=formato(objRecordset2("vlaed2"),100)%>&nbsp;mg/m<sup>2</sup></td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("vlaec1")) and objRecordset2("vlaec1")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">VLA-EC:</td>
				<td class="campo" valign="middle"><%=formato(objRecordset2("vlaec1"),100)%>&nbsp;ppm</td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("vlaec2")) and objRecordset2("vlaec2")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">VLA-EC:</td>
				<td class="campo" valign="middle"><%=formato(objRecordset2("vlaec2"),100)%>&nbsp;mg/m<sup>2</sup></td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("notas")) and objRecordset2("notas")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Notas:</td>
				<td class="celdaabajo" valign="middle"><%=formato(objRecordset2("notas"),300)%></td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>

				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 16: Valores límite (VLB)
				sql2 = "SELECT vlaed1,vlaed2,notas FROM RQ_SUSTANCIAS_VL2 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto16 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">VALORES LÍMITE AMBIENTALES CANCERÍGENOS</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("vlaed1")) and objRecordset2("vlaed1")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">VLA-ED:</td>
				<td class="campo" valign="middle"><%=formato(objRecordset2("vlaed1"),100)%>&nbsp;ppm</td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("vlaed2")) and objRecordset2("vlaed2")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">VLA-ED:</td>
				<td class="campo" valign="middle"><%=formato(objRecordset2("vlaed2"),100)%>&nbsp;mg/m<sup>2</sup></td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("notas")) and objRecordset2("notas")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Notas:</td>
				<td class="celdaabajo" valign="middle"><%=formato(objRecordset2("notas"),300)%></td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 16b: Valores límite (VLB)
				sql2 = "SELECT ib,vlb,MOMENTO_MUESTREO,notas FROM RQ_SUSTANCIAS_VL3 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto16b = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">VALORES LÍMITE BIOLÓGICOS</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("ib")) and objRecordset2("ib")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Indicador Biológico:</td>
				<td class="campo" valign="middle"><%=formato(objRecordset2("ib"),100)%>&nbsp;</td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("vlb")) and objRecordset2("vlb")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">VLB:</td>
				<td class="campo" valign="middle"><%=formato(objRecordset2("vlb"),100)%>&nbsp;</td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("MOMENTO_MUESTREO")) and objRecordset2("MOMENTO_MUESTREO")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Momento Muestreo:</td>
				<td class="campo" valign="middle"><%=formato(objRecordset2("MOMENTO_MUESTREO"),100)%>&nbsp;</td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("notas")) and objRecordset2("notas")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Notas:</td>
				<td class="celdaabajo" valign="middle"><%=formato(objRecordset2("notas"),300)%></td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>
				<div style="height:3pt"></div>
				<% end if %>
				
<% 	'-- 17: Enfermedades profesionales
				sql3 = "SELECT DISTINCT RQ_ENF_PROF.idenfprof FROM RQ_SUST_ENF LEFT JOIN RQ_ENF_PROF ON RQ_SUST_ENF.enf_prof=RQ_ENF_PROF.idenfprof LEFT JOIN RQ_SUSTANCIAS ON RQ_SUST_ENF.sustancia=RQ_SUSTANCIAS.id WHERE RQ_SUSTANCIAS.cas='"&CAS_actual&"'"
				set objRecordset3 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset3.Open sql3,objConnection,adOpenKeyset
		   	   	if not objRecordset3.eof then
		   	   		texto17 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">ENFERMEDADES PROFESIONALES RELACIONADAS (borrador)</td></tr>
				<% do while not objRecordset3.eof %>
				<% 	if objRecordset3("idenfprof")<>"" then
						sql2 = "SELECT d1,d2,d3 FROM RQ_ENF_PROF WHERE idenfprof="&objRecordset3("idenfprof")
		   	   			set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   			objRecordset2.Open sql2,objConnection,adOpenKeyset %>
				<% if not isnull(objRecordset2("d1")) and objRecordset2("d1")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Grupo:</td>
				<td class="campo" valign="middle"><b><%=objRecordset2("d1")%></b></td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("d2")) and objRecordset2("d2")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Relación indicativa de síntomas y patologías relacionadas con el agente:</td>
				<td class="campo" valign="top"><%=replace(objRecordset2("d2"),chr(13),"<br>")%></td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("d3")) and objRecordset2("d3")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Principales actividades capaces de producir enfermedades relacionadas con el agente:</td>
				<td class="celdaabajo" valign="top"><%=replace(objRecordset2("d3"),chr(13),"<br>")%></td></tr>
				<% end if %>
					<% end if %>
				<% objRecordset3.movenext
				   loop %>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<!--fin normativa salud laboral-->
<% end if %>
<!--normativa ambiental-->
				<div style="height:3pt"></div>

		                <table width="100%" cellpadding=5><tr><td><a name="normativama"></a><img src="imagenes/risctox05.gif" alt="normativa ambiental"></td>
		                <td align="right"></td></tr></table>

<% 	'-- 0c: Residuos peligrosos
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS WHERE ((RD<>'' OR isnull(RD,'nulo')<>'nulo') AND (cas='"&CAS_actual&"'))"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto0c = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">RESIDUOS PELIGROSOS</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>
<% if id_ecogente=179 then %>
<% 	'-- 0d: Vertidos
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS WHERE ((RD<>'' OR isnull(RD,'nulo')<>'nulo') AND (cas='"&CAS_actual&"'))"
				
				sql2 = "SELECT CAS FROM RQ_SUSTANCIAS WHERE ((RD<>'' AND isnull(RD,'nulo')<>'nulo') AND (cas='"&CAS_actual&"'))"
				sql2 = sql2 & " UNION SELECT CAS FROM RQ_SUSTANCIAS_CYM WHERE cas='"&CAS_actual&"'"
				sql2 = sql2 & " UNION SELECT CAS FROM RQ_SUSTANCIAS_CYM2 WHERE cas='"&CAS_actual&"'"
				sql2 = sql2 & " UNION SELECT CAS FROM RQ_SUSTANCIAS_CYM3 WHERE cas='"&CAS_actual&"'"
				sql2 = sql2 & " UNION SELECT CAS FROM RQ_SUSTANCIAS_DIS WHERE cas='"&CAS_actual&"'"
				sql2 = sql2 & " UNION SELECT CAS FROM RQ_SUSTANCIAS_PYB WHERE cas='"&CAS_actual&"'"
				sql2 = sql2 & " UNION SELECT CAS FROM RQ_SUSTANCIAS_TAC WHERE cas='"&CAS_actual&"'"
				sql2 = sql2 & " UNION SELECT CAS FROM RQ_SUSTANCIAS_TAC2 WHERE cas='"&CAS_actual&"'"
				
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto0d = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">VERTIDOS</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>
<% end if %>
<% 	'-- 18: Emisiones
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS_EMI WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto18 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">EMISIONES</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 19: COV
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS_COV WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto19 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">COV</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 20: LPCIC
				sql2 = "SELECT atmosfera,agua FROM RQ_SUSTANCIAS_LPC WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto20 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">LPCIC</td></tr>
				<% 'do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("atmosfera")) and objRecordset2("atmosfera")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top" width="50%">Atmósfera:</td>
				<td class="campo" valign="middle"><% if objRecordset2("atmosfera")="X" then response.write "SÍ" else response.write "NO"%></td></tr>
				<% end if %>
				<% if not isnull(objRecordset2("agua")) and objRecordset2("agua")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Agua:</td>
				<td class="celdaabajo" valign="middle"><% if objRecordset2("agua")="X" then response.write "SÍ" else response.write "NO"%></td></tr>
				<% end if %>
				<% 'objRecordset2.movenext
				   'loop %>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 0e: Accidentes mayores
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS WHERE ((RD<>'' OR isnull(RD,'nulo')<>'nulo') AND (simbolo LIKE '%T%' OR simbolo LIKE '%O%' "
				for i=1 to 11
					sql2 = sql2 &" OR clasific"&i&" LIKE '%R2' OR clasific"&i&" LIKE '%R3' OR clasific"&i&" LIKE '%R2-%' OR clasific"&i&" LIKE '%R3-%' OR clasific"&i&" LIKE '%10%' OR clasific"&i&" LIKE '%11%' OR clasific"&i&" LIKE '%12%' OR clasific"&i&" LIKE '%17%' OR clasific"&i&" LIKE '%50%' OR clasific"&i&" LIKE '%51%' " 
				next
				sql2 = sql2 & ") AND (cas='"&CAS_actual&"'))"
				
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto0e = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">ACCIDENTES GRAVES</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<!--fin normativa ambiental-->
<!--observaciones-->

<% 	'-- 21: Efectos sobre la salud y/o órganos afectados
				sql2 = "SELECT * FROM RISCTOX_SUSTANCIAS2 WHERE cas='"&CAS_actual&"'"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto21 = "si"
		   	   		cardiocirculatorio = objRecordset2("cardiocirculatorio")
					rinyon = objRecordset2("rinyon")
					respiratorio = objRecordset2("respiratorio")
					reproductivo = objRecordset2("reproductivo")
					piel_sentidos = objRecordset2("piel_sentidos")
					neuro_toxicos = objRecordset2("neuro_toxicos")
					musculo_esqueletico = objRecordset2("musculo_esqueletico")
					sistema_inmunitario = objRecordset2("sistema_inmunitario")
					higado_gastrointestinal = objRecordset2("higado_gastrointestinal")
					sistema_endocrino = objRecordset2("sistema_endocrino")
					embrion = objRecordset2("embrion")
					if cardiocirculatorio=1 or rinyon=1 or respiratorio=1 or reproductivo=1 or piel_sentidos=1 or neuro_toxicos=1 or musculo_esqueletico=1 or sistema_inmunitario=1 or higado_gastrointestinal=1 or sistema_endocrino=1 or embrion=1 then  %>
				<div style="height:3pt"></div>

				<table width="100%" cellpadding=5><tr><td><a name="observaciones"></a><img src="imagenes/risctox06.gif" alt="observaciones"></td>
		                <td align="right"></td></tr></table>

				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">EFECTOS SOBRE LA SALUD Y/O ÓRGANOS AFECTADOS</td></tr>
				<tr>
		   	   	<td class="texto" width="20%">&nbsp;</td>
				<td class="texto" valign="middle">
					<% if cardiocirculatorio=1 then response.write "Cardiocirculatorio"&"<br>" %>
					<% if rinyon=1 then response.write "Riñ&oacute;n"&"<br>" %>
					<% if respiratorio=1 then response.write "Respiratorio"&"<br>" %>
					<% if reproductivo=1 then response.write "Reproductivo"&"<br>" %>
					<% if piel_sentidos=1 then response.write "Piel y sentidos"&"<br>" %>
					<% if neuro_toxicos=1 then response.write "Neuro-tóxicos"&"<br>" %>
					<% if musculo_esqueletico=1 then response.write "Músculo esquelético"&"<br>" %>
					<% if sistema_inmunitario=1 then response.write "Sistema inmunitario"&"<br>" %>
					<% if higado_gastrointestinal=1 then response.write "Hígado-gastrointestinal"&"<br>" %>
					<% if embrion=1 then response.write "Embri&oacute;n"&"<br>" %>
				</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>
				<% end if %>
				
<!--fin observaciones-->
<!--sectores-->
<% 	'-- 22: Sectores
				sql2 = "SELECT DISTINCT RISCTOX_VALORES.desc1 FROM RISCTOX_VALORES LEFT JOIN RISCTOX_CLASIF2 ON RISCTOX_CLASIF2.id_sector = RISCTOX_VALORES.valor LEFT JOIN RISCTOX_SUSTANCIAS2 ON RISCTOX_SUSTANCIAS2.id=RISCTOX_CLASIF2.id_sustancia "
				sql2 = sql2 & "WHERE RISCTOX_SUSTANCIAS2.cas = '" & CAS_actual & "' ORDER BY RISCTOX_VALORES.desc1"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto22 = "si" %>
				<div style="height:3pt"></div>
				<table width="100%" cellpadding=5><tr><td><a name="sectores"></a><img src="imagenes/risctox07.gif" alt="sectores"></td>
		                <td align="right"></td></tr></table>

				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">SECTORES</td></tr>
				<tr>
		   	   	<td class="texto" width="5%">&nbsp;</td>
				<td class="texto" valign="middle">
				<% do while not objRecordset2.eof
					response.write ucase(mid(objRecordset2("desc1"),1,1))&mid(objRecordset2("desc1"),2,500)&"<br>"
				  	objRecordset2.movenext
				   loop %>
				</td>
				</tr>
				</table>

				<div style="height:3pt"></div>
				
				<% end if %>
<!--fin sectores-->
<!--alternativas-->
<% 	'-- 23: Alternativas
				
				sql2 = "SELECT RQ_ALTERNATIVAS.alternativa,RQ_ALTERNATIVAS.idalternativa FROM RQ_ALTERNATIVAS LEFT JOIN RQ_ALTERNATIVAS_RELACIONES ON RQ_ALTERNATIVAS.idalternativa=RQ_ALTERNATIVAS_RELACIONES.idalternativa LEFT JOIN RQ_SUSTANCIAS ON RQ_ALTERNATIVAS_RELACIONES.id_relacion=RQ_SUSTANCIAS.id WHERE RQ_ALTERNATIVAS_RELACIONES.tabla_relacion='RQ_SUSTANCIAS' AND RQ_SUSTANCIAS.cas='"&CAS_actual&"' ORDER BY RQ_ALTERNATIVAS.alternativa"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto23 = "si" %>
				<br>&nbsp;
				<table width="100%" cellpadding=5><tr><td><a name="alternativas"></a><img src="imagenes/risctox08.gif" alt="Alternativas"></td>
		                <td align="right"></td></tr></table>

				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">ALTERNATIVAS</td></tr>
				<% do while not objRecordset2.eof  %>
		   	   	<tr>
		   	   	<td class="texto" valign="middle"><%=ucase(objRecordset2("alternativa"))%></td>
				</tr>
				<% objRecordset2.movenext
				   loop %>
				</table>
				<div style="height:3pt"></div>
				<% end if %>
<!--fin alternativas-->
<% end if %>
<p align=center class="campo">www.ecoinformas.com</p>

</td></tr></table>
</body>
</html>
