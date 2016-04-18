<% response.redirect "dn_risctox_buscador.asp" %>
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
	
	'numeracion = "AICCA"
	idpagina = 627	'--- página Ficha sustancia, sólo para registrar estadísticas
	sql = "SELECT numeracion FROM WEBISTAS_PAGINAS WHERE idpagina="&idpagina
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	numeracion = objRecordset("numeracion")
	
	'usuario = request.cookies("webistas")
	'if cstr(usuario)="" then 
	'	sql = "SELECT max(idgente) as ultimousuario FROM WEBISTAS_VISITAS"
	'	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	'	set objRecordset = OBJConnection.Execute(sql)
	'	usuario = clng(objrecordset("ultimousuario"))+1
	'	Response.Cookies("webistas") = usuario
	'	Response.Cookies("webistas").Expires = #1/1/2010#
	'end if

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

	FUNCTION ayuda(x)
		response.write "<a onclick=window.open('ver_definicion.asp?id="&x&"','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'><img src='imagenes/ayuda.gif' width=14 height=14></a>&nbsp;"
	END FUNCTION
	
	FUNCTION lista(x)
		response.write "<a href='index.asp?idpagina="&x&"'><img src='imagenes/ayuda.gif' width=14 height=14 border=0 align=absmiddle></a>&nbsp;"
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

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: Base de datos de sustancias tóxicas y peligrosas RISCTOX</title>
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
function enviar() 
{

		if  (document.form_buscar.buscar.value == "aaa") 
			{ alert("Es necesario escribir una parte del nombre o sinónimo de la sustancia o su número CAS, CE o RD"); }
			else
			{ document.form_buscar.submit(); }
}

function mostrar(cual)
{
	eval("oculto"+cual+".style.visibility = 'hidden'");
	eval("oculto"+cual+".style.display = 'none'");
	eval("visible"+cual+".style.visibility = 'visible'");
	eval("visible"+cual+".style.display = 'block'");
}

function ocultar(cual)
{
	eval("oculto"+cual+".style.visibility = 'visible'");
	eval("oculto"+cual+".style.display = 'block'");
	eval("visible"+cual+".style.visibility = 'hidden'");
	eval("visible"+cual+".style.display = 'none'");
}

// -->
</SCRIPT>
</head>
<body>
<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
			<div id="encabezado_nuevo3">
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
			<div id="menusup3">
			<table border="0" cellspacing="5" cellpadding="0">
            			<tr class="textmenusup">
<%              				sql = "SELECT titulo,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion LIKE 'AIC%' AND len(numeracion)=4 ORDER BY numeracion"
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
			<div class="textsubmenu" id="submenusup3">
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
			

			<div id="texto">
			
				<div class="texto">
             				<% if len(numeracion)>3 then
              				     response.write "<br><p class=campo>Est&aacute;s en: "
              				     for i=1 to len(numeracion)-3
              				   	sql = "SELECT titulo,idpagina FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion='"&mid(numeracion,1,2+i)&"'" 
			   	   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	           	set objRecordset = OBJConnection.Execute(sql)
		   	   	           	if not(objRecordset.eof) then response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&">"&objRecordset("titulo")&"</a>&nbsp;&gt;&nbsp;"
              				     next
              				     response.write "<a href=risctox1.asp>Inicio</a>&nbsp;&gt;&nbsp;Ficha de la sustancia: "&request("nombre")&"</p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              			   end if %>
				
				<table width="100%"><tr>
				<td class=titulo3>RISCTOX: Ficha de la sustancia</td>
				<td align="right"><table><tr>
				<td align="right" valign="middle">Ir a:</td><td align="right"><select name="marcador" class=subtitulo3 onchange=location.href="#"+this.value>
						<option value="identificacion">Identificación de la sustancia</option>
						<option value="riesgossalud">Riesgos específicos para la salud</option>
						<option value="riesgosma">Riesgos específicos para el medioambiente</option>
						<% if id_ecogente=179 then %>
						<option value="normativasalud">Normativa salud laboral</option>
						<% end if %>
						<option value="normativama">Normativa medioambiental</option>
						<option value="observaciones">Observaciones</option>
						<option value="sectores">Sectores</option>
						<option value="alternativas">Alternativas</option>
						</select>
				</td></tr></table></td></tr></table>
				<br>&nbsp;						
<!--identificación-->
				<div id="ficha">
		                <table width="100%" cellpadding=5><tr><td><a name="identificacion"></a><img src="imagenes/risctox01.gif" alt="identificación de la sustancia" width="255" height="32" /></td>
		                <td align="right"><a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a></td></tr></table>

<%							
				CAS_actual = trim(request("CAS"))
				sql = "SELECT * FROM RQ_SUSTANCIAS WHERE CAS='"&CAS_actual&"'"
				set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset.Open sql,objConnection,adOpenKeyset
		   	   	if not objRecordset.eof then %>
<%	'-- 1: DATOS SUSTANCIA											%>
				
				<div id="oculto1" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		   	   	<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
			   	   	<a onclick="mostrar('1')" style="text-decoration:none;cursor:hand"><font class="titulo3">SUSTANCIA</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('1')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
			   	</td></tr>
		   	   	</table>
		   	   	</div>
		   	   	
		   	   	<div id="visible1" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		   	   	<tr><td class="celdaabajo" colspan="2" align="center">
		   	   	<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
			   	   	<a onclick="ocultar('1')" style="text-decoration:none;cursor:hand"><font class="titulo3">SUSTANCIA</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('1')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
		   	   	<tr><td class="subtitulo3" align="right" valign="top"><% call ayuda(82) %>Nombre:</td><td class="texto" valign="middle"><b><%=formato(objRecordset("nombre"),300)%></b></td></tr>
		   	   	<% if not isnull(objRecordset("sinonimos")) and objRecordset("sinonimos")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top"><% call ayuda(83) %>Sinónimos:</td><td class="texto" valign="middle"><b><%=formato(objRecordset("sinonimos"),300)%></b></td></tr>
		   	   	<% end if %>
		   	   	<% if not isnull(objRecordset("cas")) and objRecordset("cas")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top"><% call ayuda(84) %>CAS:</td><td class="texto" valign="middle"><% if instr(objRecordset("cas"),"-XX-")=0 then response.write formato(objRecordset("cas"),300)%></td></tr>
		   	   	<% end if %>
		   	   	<% if not isnull(objRecordset("cee")) and objRecordset("cee")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top"><% call ayuda(85) %>Nº C.E./EINECS:</td><td class="texto" valign="middle"><%=formato(objRecordset("cee"),300)%></td></tr>
		   	   	<% end if %>
		   	   	<% if not isnull(objRecordset("rd")) and objRecordset("rd")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top" nowrap><% call ayuda(86) %>R.D. 363/1995:</td><td class="texto" valign="middle"><%=formato(objRecordset("rd"),300)%></td></tr>
				<% else %>
				<tr><td class="texto" valign="middle" colspan="2" align="center">Sustancia no incluida en el Anexo I del RD 363/1995.<br>Es responsabilidad del fabricante de la sustancia o preparado asignarle las Frases R y S</td></tr>
				<% end if %>
				</table>
				</div>
				
				<div style="height:3pt"></div>
<%	'-- 2: CLASIFICACIÓN 										  %>

				<div id="oculto2" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call ayuda(87) %><a onclick="mostrar('2')" style="text-decoration:none;cursor:hand"><font class="titulo3">CLASIFICACIÓN</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('2')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible2" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call ayuda(87) %><a onclick="ocultar('2')" style="text-decoration:none;cursor:hand"><font class="titulo3">CLASIFICACIÓN</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('2')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Frases R (clasificación <%=i%>):</td><td class="texto" valign="middle"><%=objRecordset(campo)%>&nbsp;
		   	   	<a onclick="window.open('busca_frases_r.asp?id=<%=objRecordset(campo)%>','fr','width=300,height=200,scrollbars=yes,resizable=yes')" style="text-decoration:none;cursor:hand"><img src="imagenes/interrogacion.gif" border="0" align="absmiddle" alt="busca Frases R"></a></td></tr>
		   	   	<% 	end if %>
		   	   	<% next %>
		   	   	<% if not isnull(objRecordset("frases_s")) and objRecordset("frases_s")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Frases S:</td><td class="texto" valign="middle"><%=objRecordset("frases_s")%>&nbsp;
		   	   	<a onclick="window.open('busca_frases_s.asp?id=<%=objRecordset("frases_s")%>','fr','width=300,height=200,scrollbars=yes,resizable=yes')" style="text-decoration:none;cursor:hand"><img src="imagenes/interrogacion.gif" border="0" align="absmiddle" alt="busca Frases S"></a></td></tr>
				<% end if %>
				</table>
				</div>
				
				<div style="height:3pt"></div>

<%	'-- 3: ETIQUETADO
				   eticonc= ""
				   for i=1 to 11
					if not isnull(objrecordset("eticonc"&cstr(i))) then eticonc = eticonc & objrecordset("eticonc"&cstr(i))
				   next
				   if eticonc<>"" then %>
				<div id="oculto3" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call ayuda(88) %><a onclick="mostrar('3')" style="text-decoration:none;cursor:hand"><font class="titulo3">ETIQUETADO</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('3')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible3" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call ayuda(88) %><a onclick="ocultar('3')" style="text-decoration:none;cursor:hand"><font class="titulo3">ETIQUETADO</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('3')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				<% for i=1 to 11 
					campo = "eticonc"&cstr(i)
					campo2 = "conc"&cstr(i)
					if not isnull(objRecordset(campo)) and objRecordset(campo)<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top" nowrap>Concentración <%=i%>:</td>
		   	   	    <td class="texto" valign="middle"><%=objRecordset(campo2)%></td>
		   	   	    <td class="subtitulo3" align="right" valign="top" nowrap>Etiqueta <%=i%>:</td>
		   	   	    <td class="texto" valign="middle"><%=objRecordset(campo)%>&nbsp;
		   	   	    <a onclick="window.open('busca_frases_r.asp?id=<%=objRecordset(campo)%>','fr','width=300,height=200,scrollbars=yes,resizable=yes')" style="text-decoration:none;cursor:hand"><img src="imagenes/interrogacion.gif" border="0" align="absmiddle" alt="busca Frases R"></a>
		   	   	    </td>
		   	   	</tr>
		   	   	<% 	end if %>
		   	   	<% next %>
				</table>
				</div>
				<div style="height:3pt"></div>
				<% end if %>
<% else %>
<%	'-- 1 danés: DATOS SUSTANCIA											%>
<%				sql2 = "SELECT * FROM RQ_SUSTANCIAS_DANESAS WHERE CAS='"&CAS_actual&"'"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then %>
				
				<div id="oculto1" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		   	   	<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
			   	   	<a onclick="mostrar('1')" style="text-decoration:none;cursor:hand"><font class="titulo3">SUSTANCIA</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('1')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
			   	</td></tr>
		   	   	</table>
		   	   	</div>
		   	   	
		   	   	<div id="visible1" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		   	   	<tr><td class="celdaabajo" colspan="2" align="center">
		   	   	<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
			   	   	<a onclick="ocultar('1')" style="text-decoration:none;cursor:hand"><font class="titulo3">SUSTANCIA</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('1')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
		   	   	<tr><td class="subtitulo3" align="right" valign="top"><% call ayuda(82) %>Nombre:</td><td class="texto" valign="middle"><b><%=formato(request("nombre"),300)%></b></td></tr>
		   	   	<% if not isnull(objRecordset2("cas")) and objRecordset2("cas")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top"><% call ayuda(84) %>CAS:</td><td class="texto" valign="middle"><%=formato(objRecordset2("cas"),300)%></td></tr>
		   	   	<% end if %>
		   	   	<% if not isnull(objRecordset2("einecs")) and objRecordset2("einecs")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top"><% call ayuda(85) %>Nº C.E./EINECS</td><td class="texto" valign="middle"><%=formato(objRecordset2("einecs"),300)%></td></tr>
		   	   	<% end if %>
		   	   	<tr><td class="texto" valign="middle" colspan="2" align="center">Sustancia no incluida en el Anexo I del RD 363/1995.<br>Es responsabilidad del fabricante de la sustancia o preparado asignarle las Frases R y S</td></tr>
				</table>
				</div>
				<div style="height:3pt"></div>
<%	'-- 2 danés: CLASIFICACIÓN 										  %>

				<div id="oculto2" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<a onclick="mostrar('2')" style="text-decoration:none;cursor:hand"><font class="titulo3">CLASIFICACIÓN</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('2')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible2" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<a onclick="ocultar('2')" style="text-decoration:none;cursor:hand"><font class="titulo3">CLASIFICACIÓN</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('2')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				<% if not isnull(objRecordset2("frases_r")) and objRecordset2("frases_r")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Frases R (recomendadas por la Agencia de Medio Ambiente de Dinamarca):</td><td class="texto" valign="middle"><%=objRecordset2("frases_r")%>&nbsp;
		   	   	<a onclick="window.open('busca_frases_r.asp?id=<%=objRecordset2("frases_r")%>','fr','width=300,height=200,scrollbars=yes,resizable=yes')" style="text-decoration:none;cursor:hand"><img src="imagenes/interrogacion.gif" border="0" align="absmiddle" alt="busca Frases R"></a></td></tr>
		   	   	<% end if %>
				</table>
				</div>
				
				<div style="height:3pt"></div>

				<% else %>

<%	'-- 1 no está en RD ni en la lista danesa: DATOS SUSTANCIA											%>
				<div id="oculto1" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		   	   	<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
			   	   	<a onclick="mostrar('1')" style="text-decoration:none;cursor:hand"><font class="titulo3">SUSTANCIA</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('1')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
			   	</td></tr>
		   	   	</table>
		   	   	</div>
		   	   	
		   	   	<div id="visible1" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
		   	   	<tr><td class="celdaabajo" colspan="2" align="center">
		   	   	<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
			   	   	<a onclick="ocultar('1')" style="text-decoration:none;cursor:hand"><font class="titulo3">SUSTANCIA</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('1')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
		   	   	<tr><td class="subtitulo3" align="right" valign="top"><% call ayuda(82) %>Nombre:</td><td class="texto" valign="middle"><b><%=formato(request("nombre"),300)%></b></td></tr>
		   	   	<tr><td class="subtitulo3" align="right" valign="top"><% call ayuda(84) %>CAS:</td><td class="texto" valign="middle"><%=formato(request("cas"),300)%></td></tr>
				<tr><td class="texto" valign="middle" colspan="2" align="center">Sustancia no incluida en el Anexo I del RD 363/1995.<br>Es responsabilidad del fabricante de la sustancia o preparado asignarle las Frases R y S</td></tr>
				</table>
				</div>
				<div style="height:3pt"></div>

				<% end if %>
<% end if %>
				</div>
				<div style="height:3pt"></div>
<!--fin de identificación-->
				<p align="center"><input type="button" class="boton" value="imprimir ficha identificación" onclick="window.open('imprimir_ficha.asp?ficha=identificacion&cas=<%=CAS_actual%>&nombre=<%=request("nombre")%>','ficha','width=300,height=300,resizable=yes,scrollbars=yes')"></p>

<!--riesgos para la salud-->				
				<div id="ficha">
		                <table width="100%" cellpadding=5><tr><td><a name="riesgossalud"></a><img src="imagenes/risctox02.gif" alt="riesgos específicos para la salud"></td>
		                <td align="right"><a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a></td></tr></table>

<% 	'-- 4: Cancerígenas y mutágenas (según RD 363/1995)
				sql2 = "SELECT C,M,notas FROM RQ_SUSTANCIAS_CYM WHERE cas='"&CAS_actual&"' GROUP BY C,M,notas"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then 
		   	   		texto4="si" %>
				<div id="oculto4" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(607)%><a onclick="mostrar('4')" style="text-decoration:none;cursor:hand"><font class="titulo3">CANCERÍGENA Y MUTÁGENA</font> (según RD 363/1995)</a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('4')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible4" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(607)%><a onclick="ocultar('4')" style="text-decoration:none;cursor:hand"><font class="titulo3">CANCERÍGENA Y MUTÁGENA</font> (según RD 363/1995)</a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('4')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
		   	   	<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("C")) and objRecordset2("C")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Cancerígeno:</td><td class="texto" valign="middle">
		   	   		<% if objRecordset2("c")="C1" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=1','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>C1</a>"
					   end if
					   if objRecordset2("c")="C2" then
					   	response.write "<a onclick=window.open('ver_definicion.asp?id=2','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>C2</a>"
					   end if 
					   if objRecordset2("c")="C3" then
					   	response.write "<a onclick=window.open('ver_definicion.asp?id=81','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>C3</a>"
					   end if %>
		   	   	</td></tr>
				<% end if %>
		   	   	<% if not isnull(objRecordset2("M")) and objRecordset2("M")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Mutágeno:</td><td class="texto" valign="middle">
		   	   		<% if objRecordset2("M")="M2" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=3','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>M2</a>"
					   end if
					   if objRecordset2("M")="M3" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=91','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>M3</a>"
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
				</div>
				<div style="height:3pt"></div>
				<% end if %>
				
<% 	'-- 5: Cancerígenas y mutágenas (según IARC)
				sql2 = "SELECT grupo,volumen FROM RQ_SUSTANCIAS_CYM2 WHERE cas='"&CAS_actual&"' GROUP BY grupo,volumen"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto5 = "si" %>
				<div id="oculto5" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(607)%><a onclick="mostrar('5')" style="text-decoration:none;cursor:hand"><font class="titulo3">CANCERÍGENA Y MUTÁGENA</font> (según IARC)</a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('5')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible5" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(607)%><a onclick="ocultar('5')" style="text-decoration:none;cursor:hand"><font class="titulo3">CANCERÍGENA Y MUTÁGENA</font> (según IARC)</a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('5')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
				<div style="height:3pt"></div>
				<% end if %>
				

<% 	'-- 6: Cancerígenas y mutágenas (según otras fuentes)
				sql2 = "SELECT fuente FROM RQ_SUSTANCIAS_CYM3 WHERE cas='"&CAS_actual&"'"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then 
		   	   		texto6 = "si" %>
				<div id="oculto6" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(607)%><a onclick="mostrar('6')" style="text-decoration:none;cursor:hand"><font class="titulo3">CANCERÍGENA Y MUTÁGENA</font> (según otras fuentes)</a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('6')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible6" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(607)%><a onclick="ocultar('6')" style="text-decoration:none;cursor:hand"><font class="titulo3">CANCERÍGENA Y MUTÁGENA</font> (según otras fuentes)</a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('6')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
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
					<% call lista(609)%><font class="titulo3">TÓXICO PARA LA REPRODUCCIÓN</font>
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
				<div id="oculto7" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(610)%><a onclick="mostrar('7')" style="text-decoration:none;cursor:hand"><font class="titulo3">DISRUPTOR ENDOCRINO</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('7')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible7" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(610)%><a onclick="ocultar('7')" style="text-decoration:none;cursor:hand"><font class="titulo3">DISRUPTOR ENDOCRINO</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('7')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
				<div style="height:3pt"></div>
				<% end if %>

<% if id_ecogente=179 then %>
<% 	'-- 8: Neurotóxicos
				sql2 = "SELECT efecto,nivel,fuente FROM RQ_SUSTANCIAS_NEU WHERE cas='"&CAS_actual&"'"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto8 = "si" %>
				<div id="oculto8" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(611)%><a onclick="mostrar('8')" style="text-decoration:none;cursor:hand"><font class="titulo3">NEUROTÓXICO</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('8')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible8" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(611)%><a onclick="ocultar('8')" style="text-decoration:none;cursor:hand"><font class="titulo3">NEUROTÓXICO</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('8')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
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
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(612)%><font class="titulo3">SENSIBILIZANTE</font>
					</td><td width="20%" align="right">
					&nbsp;
					</td></tr></table>
				</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>
<% end if %>
<% if 1=0 and texto4<>"si" and texto5<>"si" and texto6<>"si" and texto7<>"si" and texto8<>"si" and texto0<>"si" and texto0b<>"si" then %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="texto" align="center">No se conocen riesgos para la salud</td></tr>
				</table>
				<br>&nbsp;
<% end if %>
				</div>
<!--fin de riesgos para la salud-->
<!--riesgos para el medioambiente-->
				<div style="height:3pt"></div>
				<div id="ficha">
		                <table width="100%" cellpadding=5><tr><td><a name="riesgosma"></a><img src="imagenes/risctox03.gif" alt="riesgos específicos para el medioambiente"></td>
		                <td align="right"><a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a></td></tr></table>

<% 	'-- 9: Persistencia y bioacumulación
				sql2 = "SELECT enlace,url FROM RQ_SUSTANCIAS_PYB WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto9 = "si" %>
				<div id="oculto9" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(613)%><a onclick="mostrar('9')" style="text-decoration:none;cursor:hand"><font class="titulo3">TÓXICAS, PERSISTENTES Y BIOACUMULATIVAS</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('9')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible9" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(613)%><a onclick="ocultar('9')" style="text-decoration:none;cursor:hand"><font class="titulo3">TÓXICAS, PERSISTENTES Y BIOACUMULATIVAS</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('9')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("url")) and objRecordset2("url")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Más información (en inglés):</td><td class="texto" valign="middle">
					 <a href="<%=lcase(objRecordset2("url"))%>" target="_blank"><%=mid(lcase(objRecordset2("enlace")),1,100)%></a>&nbsp;
		   	   	</td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>
				</div>
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
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(614)%><font class="titulo3">TOXICIDAD ACUÁTICA</font> (según directiva de aguas)
					</td><td width="20%" align="right">
					&nbsp;
					</td></tr></table>
				</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 11: Toxicidad acuática (Peligrosas agua Alemania)
				sql2 = "SELECT campo5 FROM RQ_SUSTANCIAS_TAC2 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto11 = "si" %>
				<div id="oculto11" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(614)%><a onclick="mostrar('11')" style="text-decoration:none;cursor:hand"><font class="titulo3">TOXICIDAD ACUÁTICA</font> (Peligrosas agua Alemania)</a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('11')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible11" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(614)%><a onclick="ocultar('11')" style="text-decoration:none;cursor:hand"><font class="titulo3">TOXICIDAD ACUÁTICA</font> (Peligrosas agua Alemania)</a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('11')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
				<div style="height:3pt"></div>
				<% end if %>
<% end if %>
<% 	'-- 12: Daño a la atmósfera (Capa de Ozono)
				sql2 = "SELECT nombre2 FROM RQ_SUSTANCIAS_DAT WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto12 = "si" %>
				<div id="oculto12" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(615)%><a onclick="mostrar('12')" style="text-decoration:none;cursor:hand"><font class="titulo3">DAÑO A LA ATMÓSFERA</font> (Capa de Ozono)</a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('12')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible12" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(615)%><a onclick="ocultar('12')" style="text-decoration:none;cursor:hand"><font class="titulo3">DAÑO A LA ATMÓSFERA</font> (Capa de Ozono)</a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('12')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				<% do while not objRecordset2.eof %>
		   	   	<% if not isnull(objRecordset2("nombre2")) and objRecordset2("nombre2")<>"" then %>
		   	   	<tr><td class="subtitulo3" align="right" valign="top">Otro nombre:</td>
				<td class="texto" valign="middle"><%=formato(objRecordset2("nombre2"),100)%></td></tr>
				<% end if %>
				<% objRecordset2.movenext
				   loop %>
				</table>
				</div>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 13: Daño a la atmósfera (cambio climático)
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS_DAT2 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto13 = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(615)%><font class="titulo3">DAÑO A LA ATMÓSFERA</font> (cambio climático)
					</td><td width="20%" align="right">
					&nbsp;
					</td></tr></table>
				</td></tr>
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
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(615)%><font class="titulo3">DAÑO A LA ATMÓSFERA</font> (calidad del aire)
					</td><td width="20%" align="right">
					&nbsp;
					</td></tr></table>
				</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% if 1=0 and texto9<>"si" and texto10<>"si" and texto11<>"si" and texto12<>"si" and texto13<>"si" and texto14<>"si" then %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="texto" align="center">No se conocen riesgos para el medioambiente</td></tr>
				</table><br>&nbsp;
<% end if %>				

				</div>
<!--fin de riesgos para el medioambiente-->

<% if id_ecogente=179 then %>
<!--normativa salud laboral-->
				<div style="height:3pt"></div>
				<div id="ficha">
		                <table width="100%" cellpadding=5><tr><td><a name="normativasalud"></a><img src="imagenes/risctox04.gif" alt="normativa salud laboral"></td>
		                <td align="right"><a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a></td></tr></table>


<% 	'-- 15: Valores límite (VLA)
				sql2 = "SELECT vlaed1,vlaed2,vlaec1,vlaec2,notas FROM RQ_SUSTANCIAS_VL1 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto15 = "si" %>
				<div id="oculto15" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(616)%><a onclick="mostrar('15')" style="text-decoration:none;cursor:hand"><font class="titulo3">VALORES LÍMITE AMBIENTALES</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('15')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible15" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(616)%><a onclick="ocultar('15')" style="text-decoration:none;cursor:hand"><font class="titulo3">VALORES LÍMITE AMBIENTALES</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('15')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 16: Valores límite (VLB)
				sql2 = "SELECT vlaed1,vlaed2,notas FROM RQ_SUSTANCIAS_VL2 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto16 = "si" %>
				<div id="oculto16" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(616)%><a onclick="mostrar('16')" style="text-decoration:none;cursor:hand"><font class="titulo3">VALORES LÍMITE AMBIENTALES CANCERÍGENOS</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('16')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible16" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(616)%><a onclick="ocultar('16')" style="text-decoration:none;cursor:hand"><font class="titulo3">VALORES LÍMITE AMBIENTALES CANCERÍGENOS</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('16')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 16b: Valores límite (VLB)
				sql2 = "SELECT ib,vlb,MOMENTO_MUESTREO,notas FROM RQ_SUSTANCIAS_VL3 WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto16b = "si" %>
				<div id="oculto16b" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(616)%><a onclick="mostrar('16b')" style="text-decoration:none;cursor:hand"><font class="titulo3">VALORES LÍMITE BIOLÓGICOS</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('16b')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible16b" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(616)%><a onclick="ocultar('16b')" style="text-decoration:none;cursor:hand"><font class="titulo3">VALORES LÍMITE BIOLÓGICOS</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('16b')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
				<div style="height:3pt"></div>
				<% end if %>
				
<% 	'-- 17: Enfermedades profesionales
				sql3 = "SELECT DISTINCT RQ_ENF_PROF.idenfprof FROM RQ_SUST_ENF LEFT JOIN RQ_ENF_PROF ON RQ_SUST_ENF.enf_prof=RQ_ENF_PROF.idenfprof LEFT JOIN RQ_SUSTANCIAS ON RQ_SUST_ENF.cas=RQ_SUSTANCIAS.cas WHERE RQ_SUSTANCIAS.cas='"&CAS_actual&"'"
				set objRecordset3 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset3.Open sql3,objConnection,adOpenKeyset
		   	   	if not objRecordset3.eof then
		   	   		texto17 = "si" %>
				<div id="oculto17" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(617)%><a onclick="mostrar('17')" style="text-decoration:none;cursor:hand"><font class="titulo3">ENFERMEDADES PROFESIONALES RELACIONADAS (borrador)</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('17')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible17" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(617)%><a onclick="ocultar('17')" style="text-decoration:none;cursor:hand"><font class="titulo3">ENFERMEDADES PROFESIONALES RELACIONADAS (borrador)</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('17')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
				<div style="height:3pt"></div>
				<% end if %>

<% if 1=0 and texto15<>"si" and texto16<>"si" and texto17<>"si" then %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="texto" align="center">No se tiene información</td></tr>
				</table><br>&nbsp;
<% end if %>

				</div>
<!--fin normativa salud laboral-->
<% end if %>
<!--normativa ambiental-->
				<div style="height:3pt"></div>
				<div id="ficha">
		                <table width="100%" cellpadding=5><tr><td><a name="normativama"></a><img src="imagenes/risctox05.gif" alt="normativa ambiental"></td>
		                <td align="right"><a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a></td></tr></table>

<% 	'-- 0c: Residuos peligrosos
				sql2 = "SELECT cas FROM RQ_SUSTANCIAS WHERE ((RD<>'' OR isnull(RD,'nulo')<>'nulo') AND (cas='"&CAS_actual&"'))"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto0c = "si" %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(618)%><font class="titulo3">RESIDUOS PELIGROSOS</font>
					</td><td width="20%" align="right">
					&nbsp;
					</td></tr></table>
				</td></tr>
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
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(619)%><font class="titulo3">VERTIDOS</font>
					</td><td width="20%" align="right">
					&nbsp;
					</td></tr></table>
				</td></tr>
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
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(620)%><font class="titulo3">EMISIONES</font>
					</td><td width="20%" align="right">
					&nbsp;
					</td></tr></table>
				</td></tr>
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
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(621)%><font class="titulo3">COV</font>
					</td><td width="20%" align="right">
					&nbsp;
					</td></tr></table>
				</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% 	'-- 20: LPCIC
				sql2 = "SELECT atmosfera,agua FROM RQ_SUSTANCIAS_LPC WHERE (cas='"&CAS_actual&"')"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	   	if not objRecordset2.eof then
		   	   		texto20 = "si" %>
				<div id="oculto20" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(622)%><a onclick="mostrar('20')" style="text-decoration:none;cursor:hand"><font class="titulo3">LPCIC</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('20')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible20" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(622)%><a onclick="ocultar('20')" style="text-decoration:none;cursor:hand"><font class="titulo3">LPCIC</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('20')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
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
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call lista(623)%><font class="titulo3">ACCIDENTES GRAVES</font>
					</td><td width="20%" align="right">
					&nbsp;
					</td></tr></table>
				</td></tr>
				</table>
				<div style="height:3pt"></div>
				<% end if %>

<% if 1=0 and texto0c<>"si" and texto0d<>"si" and texto0e<>"si" and texto18<>"si" and texto19<>"si" and texto20<>"si" then %>
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="texto" align="center">No se tiene información</td></tr>
				</table><br>&nbsp;
<% end if %>

				</div>
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
				<div id="ficha">
				<table width="100%" cellpadding=5><tr><td><a name="observaciones"></a><img src="imagenes/risctox06.gif" alt="observaciones"></td>
		                <td align="right"><a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a></td></tr></table>

				<div id="oculto21" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call ayuda(89) %><a onclick="mostrar('21')" style="text-decoration:none;cursor:hand"><font class="titulo3">EFECTOS SOBRE LA SALUD Y/O ÓRGANOS AFECTADOS</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('21')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible21" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call ayuda(89) %><a onclick="ocultar('21')" style="text-decoration:none;cursor:hand"><font class="titulo3">EFECTOS SOBRE LA SALUD Y/O ÓRGANOS AFECTADOS</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('21')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
				<div style="height:3pt"></div>
				
				</div>
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
				<div id="ficha">
				<table width="100%" cellpadding=5><tr><td><a name="sectores"></a><img src="imagenes/risctox07.gif" alt="sectores"></td>
		                <td align="right"><a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a></td></tr></table>


				<div id="oculto22" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call ayuda(90) %><a onclick="mostrar('22')" style="text-decoration:none;cursor:hand"><font class="titulo3">SECTORES</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('22')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible22" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<% call ayuda(90) %><a onclick="ocultar('22')" style="text-decoration:none;cursor:hand"><font class="titulo3">SECTORES</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('22')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
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
				</div>
				<div style="height:3pt"></div>
				
				</div>
				<% end if %>
<!--fin sectores-->
<!--alternativas-->
<% 	'-- 23: Alternativas
				
				'sql2 = "SELECT RQ_ALTERNATIVAS.alternativa,RQ_ALTERNATIVAS.idalternativa FROM RQ_ALTERNATIVAS LEFT JOIN RQ_ALTERNATIVAS_RELACIONES ON RQ_ALTERNATIVAS.idalternativa=RQ_ALTERNATIVAS_RELACIONES.idalternativa LEFT JOIN RQ_SUSTANCIAS ON RQ_ALTERNATIVAS_RELACIONES.id_relacion=RQ_SUSTANCIAS.id WHERE RQ_ALTERNATIVAS_RELACIONES.tabla_relacion='RQ_SUSTANCIAS' AND RQ_SUSTANCIAS.cas='"&CAS_actual&"' ORDER BY RQ_ALTERNATIVAS.alternativa"
				sql2 = "SELECT RQ_ALTERNATIVAS.alternativa,RQ_ALTERNATIVAS.idalternativa FROM RQ_ALTERNATIVAS LEFT JOIN RQ_ALTERNATIVAS_SUSTANCIAS ON RQ_ALTERNATIVAS.idalternativa=RQ_ALTERNATIVAS_SUSTANCIAS.alternativa WHERE RQ_ALTERNATIVAS_SUSTANCIAS.cas='"&CAS_actual&"'"
				set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	objRecordset2.Open sql2,objConnection,adOpenKeyset
		   	if not objRecordset2.eof then
		   	   		texto23 = "si" %>
				<br>&nbsp;
				<div id="ficha">
				<table width="100%" cellpadding=5><tr><td><a name="alternativas"></a><img src="imagenes/risctox08.gif" alt="Alternativas"></td>
		                <td align="right"><a href="#"><img src="imagenes/subir.gif" border=0 alt=subir></a></td></tr></table>

				<div id="oculto23" style="overflow: auto; visibility: hidden; display: none"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="2" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<a onclick="mostrar('23')" style="text-decoration:none;cursor:hand"><font class="titulo3">ALTERNATIVAS</font></a>
					</td><td width="20%" align="right">
					<a onclick="mostrar('23')" style="text-decoration:none;cursor:hand"><img src="imagenes/desplegar.gif" alt="desplegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				</table>
				</div>
				<div id="visible23" style="overflow: auto; visibility: visible; display: block"> 
				<table class="ficharisctox" width="90%" align="center" border="0" cellpadding="4" cellspacing="0">
				<tr><td class="celdaabajo" colspan="4" align="center">
					<table cellpadding=0 cellspacing=0 width="100%"><tr><td width="80%">
					<a onclick="ocultar('23')" style="text-decoration:none;cursor:hand"><font class="titulo3">ALTERNATIVAS</font></a>
					</td><td width="20%" align="right">
					<a onclick="ocultar('23')" style="text-decoration:none;cursor:hand"><img src="imagenes/plegar.gif" alt="plegar esta tabla" border="0"></a>
					</td></tr></table>
				</td></tr>
				<% do while not objRecordset2.eof  %>
		   	   	<tr>
		   	   	<td class="texto" valign="middle"><a href="alternativa.asp?id=<%=objRecordset2("idalternativa")%>"><%=ucase(objRecordset2("alternativa"))%></a></td>
				</tr>
				<% objRecordset2.movenext
				   loop %>
				</table>
				</div>
				<div style="height:3pt"></div>
				
				</div>
				<% end if %>
<!--fin alternativas-->

				<p align="center"><input type="button" class="boton" value="imprimir ficha completa" onclick="window.open('imprimir_ficha.asp?cas=<%=CAS_actual%>&nombre=<%=request("nombre")%>','ficha','width=300,height=300,resizable=yes,scrollbars=yes')"></p>
				</div>
				</div>

				<p align=center><object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" width=530 height=100 id="anima0" align="middle"><param name="allowScriptAccess" value="sameDomain" /><param name="movie" value="http://www.istas.net/recursos/ANI/ISTAS_01078.swf" />
				<param name="quality" value="high" /><param name="wmode" value="transparent" /><param name="bgcolor" value="#ffffff" /><embed src="http://www.istas.net/recursos/ANI/ISTAS_01078.swf" quality="high" wmode="transparent" bgcolor="#ffffff" width=530 height=100 name="" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" /></object></p>
				<p>&nbsp;</p>
			</div>

     			<map name="Map3" id="Map3">
            		<area shape="rect" coords="300,18,392,80" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="393,18,539,80" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,18,694,80" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<img src="imagenes/pie3.jpg" width="708" border="0" usemap="#Map3">
      			
    			
    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>
</body>
</html>
