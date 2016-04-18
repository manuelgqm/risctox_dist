<% if cstr(request("paso"))="" then response.redirect "dn_risctox_buscador.asp" %>
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

	if cstr(request("paso"))="1" then
	if session("id_ecogente")="" then 
		session("id_ecogente") = 4
		usuario = 4
	end if
	end if
	'----- Si es restringida y no estás identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina=575"
	id_ecogente = session("id_ecogente")
	'---- ATENCIÓN: ponerlo cuando publiquemos en abierto
	
	'numeracion = "AICCA"
	idpagina = 625	'--- página Inicio, sólo para registrar estadísticas
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
		texto = replace(texto,"<t>","<font class=titulo3>")
		texto = replace(texto,"</t>","</font>")
		texto = replace(texto,"<st>","<font class=subtitulo3>")
		texto = replace(texto,"</st>","</font>")
		texto = replace(texto,"<pd>","<table width=95% align=center cellpadding=10 cellspacing=0 class=tabla><tr><td>")
		texto = replace(texto,"</pd>","</td><td valign=top align=center><img src=pd.gif></td></tr></table>")
		vistaprevia = texto
		
	END FUNCTION

	FUNCTION lista(x)
		response.write "<a href='index.asp?idpagina="&x&"'><img src='imagenes/ayuda.gif' width=14 height=14 border=0 align=absmiddle></a>&nbsp;"
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

// -->
</SCRIPT>
</head>

<body>
<% if request("paso")<>"" then 
	if request("vv")="risctox" then 
		fic="paso1_2007.swf"
	else
		fic="paso1.swf"
	end if
	%>
<div style="position:absolute; width=778px; height=400px; top:0px; z-index:3; align=center">
<table width="778" align="center"><tr><td align="center">
<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" width="778" height="714" id="paso<%=request("paso")%>" align="middle">
<param name="allowScriptAccess" value="sameDomain" />
<param name="movie" value="<%=fic %>">
<param name="quality" value="high" />
<param name="wmode" value="transparent" />
<param name="loop" value="false" />
<param name="menu" value="false" />
<embed src="<%=fic %>" width="778" height="714" quality="high" wmode="transparent" bgcolor="#ffffff" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
</object>
</td></tr></table>
</div>
<% end if %>
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
			<% if request("paso")<>"" then
					usuario_texto = "Visita virtual: paso "&request("paso")
				else
					if session("id_ecogente")<>"" then
										sql = "SELECT nombre,apellidos,sexo FROM ECOINFORMAS_GENTE WHERE idgente="&session("id_ecogente")
			   	   				set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	        set objRecordset = OBJConnection.Execute(sql)
		   	   	        if not objRecordset.eof then
		   	   	        	usuario = objRecordset("nombre")&" "&objRecordset("apellidos")
		   	   	        	usuario_sexo = "o"
		   	   	        	if objRecordset("sexo")=75 then usuario_sexo = "a"
		   	   	        	usuario_texto = "Usuari" & usuario_sexo & " identificad" & usuario_sexo & ":&nbsp;" & usuario & "&nbsp;"
		   	   	        end if
		   	  end if
		   	end if %>
			<div class="textsubmenu" id="submenusup3">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
           			<tr><td align="right"><%=usuario_texto%></td></tr>
          		</table>
			</div>

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
              				     response.write "Inicio</p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              			   end if %>
				
		            	<table width="100%" class="campo">
		            	<tr>
		            	<td class="titulo3">Base de datos de sustancias tóxicas y peligrosas RISCTOX
				
				<form name="form_buscar" action="risctox2.asp">
				<table class="tabla3" width="90%" align="center" border=0>
				  <tr>
					<td class="subtitulo3" nowrap align="left" width="160">Buscador de sustancias:</td>
					<td class="texto"><input type="text" size="30" class="campo" name="buscar">&nbsp;<input type="button" class="boton" value="buscar" onclick="enviar()"></td>
				  </tr>
				  <tr>
					<td style="font-weight: normal" class="texto" colspan="2">&nbsp;(puedes escribir una parte del nombre o sinónimo,<br> o el número CAS, CE o RD)</td>
				  </tr>
				</table>
				</form>
				
				</td>
				<%if 1=0 then %><td align="right" valign="top"><img src="imagenes/posit2.gif" alt="ESTAMOS EN PRUEBAS. DISCULPEN LOS FALLOS QUE PUEDAN SURGIR"></td><% end if %>
				</tr>
				</table>
				
				<table width="90%" align="center" class="tabla3">
				<tr><td valign="top" width="45%">
					<table width="100%" align="center">
					<tr><td class="subtitulo3"><img src="imagenes/ico_danos_sl.gif" alt="Riesgos específicos para la salud" align="absmiddle">&nbsp;Riesgos específicos para la salud</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(607)%>Cancerígenos y mutágenos:</li><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="risctox_cym.asp">Según RD 363/1995</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="risctox_cym2.asp">Según IARC</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="risctox_cym3.asp">Según otras fuentes</a>
					</td></tr>
					<% if id_ecogente=179 then %>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(609)%><a href="risctox_tpr.asp">Tóxicos para la reproducción</a></li></td></tr>
					<% end if %>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(610)%><a href="risctox_dis.asp">Disruptores endocrinos</a></li></td></tr>
					<% if id_ecogente=179 then %>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(611)%><a href="risctox_neu.asp">Neurotóxicos</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(612)%><a href="risctox_sen.asp">Sensibilizantes</a></li></td></tr>
					<% end if %>
					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td><td valign="top" width="45%">
					<table width="100%" align="center">
					<tr><td class="subtitulo3"><img src="imagenes/ico_danos_ma.gif" alt="Riesgos específicos medioambiente" align="absmiddle">&nbsp;Riesgos específicos medioambiente</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(613)%><a href="risctox_pyb.asp">Tóxicas, persistentes y bioacumulativas</a></li></td></tr>
					<% if id_ecogente=179 then %>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(614)%>Toxicidad acuática:</li><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="risctox_tac.asp">Directiva de aguas</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="risctox_tac2.asp">Peligrosas agua Alemania</a><br>
					</td></tr>
					<% end if %>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(615)%>Daño a la atmósfera:</li><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="risctox_dat.asp">Capa de Ozono</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="risctox_dat2.asp">Cambio climático</a><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="risctox_dat3.asp">Calidad del aire</a>
					</td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td></tr>
				</table>					
				
				<br>&nbsp;
				
				<table class="tabla3" width="90%" align="center">
				<tr>
				<% if id_ecogente=179 then %>
				<td valign="top" width="45%">
					<table width="100%" align="center">
					<tr><td class="subtitulo3"><img src="imagenes/ico_normativa.gif" alt="Normativa sobre salud laboral" align="absmiddle">&nbsp;Normativa sobre salud laboral</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(616)%>Límites de exposición profesional:</li>
					<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="risctox_vl1.asp">Valores Límite Ambientales</a></td></tr></table>
					<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="risctox_vl2.asp">Valores Límite Ambientales Cancerígenos</a></td></tr></table>
					<table cellpadding=0 cellspacing=0 border=0><tr><td class="campo">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class="campo"><a href="risctox_vl3.asp">Valores Límite Biológicos</a></td></tr></table>
					</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(617)%><a href="risctox_enf.asp">Enfermedades profesionales (borrador)</a></li></td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td>
				<% end if %>
				<td valign="top" width="45%">
					<table width="100%" align="center">
					<tr><td class="subtitulo3"><img src="imagenes/ico_normativa.gif" alt="Normativa ambiental" align="absmiddle">&nbsp;Normativa ambiental</td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(618)%><a href="risctox_res.asp">Residuos peligrosos</a></li></td></tr>
					<% if id_ecogente=179 then %>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(619)%><a href="risctox_ver.asp">Vertidos</a></li></td></tr>
					<% end if %>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(620)%><a href="risctox_emi.asp">Emisiones</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(621)%><a href="risctox_cov.asp">COV</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(622)%><a href="risctox_lpc.asp">LPCIC</a></li></td></tr>
					<tr><td class="texto"><li class=vineta_risctox><% call lista(623)%><a href="risctox_acm.asp">Accidentes graves</a></li></td></tr>
					<tr><td class="texto">&nbsp;</td></tr>
					</table>
				</td></tr>
				</table>					
				
				<br>&nbsp;
								
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
<% session("id_ecogente")="" %>