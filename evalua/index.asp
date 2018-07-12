<!--#include file="../../EliminaInyeccionSQL.asp"-->
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	
	
	'if request("idpagina")="962" then response.redirect "dn_auto_introduccion.asp"
	if request("idpagina")="963" then response.redirect "dn_auto_portada.asp"
	if request("idpagina")="964" then response.redirect "dn_auto_herramienta.asp"
	if request("idpagina")="575" then response.redirect "http://www.istas.net/risctox/index.asp?idpagina=575"
	
	'Sergio
	if request("idpagina") = "" then
		response.redirect "dn_auto_portada.asp"
	end if
		
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "driver={sql server};server=disoltec02;database=istas_web;UID=xip_web;PWD=***REMOVED**"

	'usuario = request.cookies("webistas")
	
	'if cstr(usuario)="" then 
	'	sql = "SELECT max(idgente) as ultimousuario FROM WEBISTAS_VISITAS"
	'	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	'	set objRecordset = OBJConnection.Execute(sql)
	'	usuario = clng(objrecordset("ultimousuario"))+1
	'	Response.Cookies("webistas") = usuario
	'	Response.Cookies("webistas").Expires = #1/1/2010#
	'end if



	idpagina = request("idpagina")
	if cstr(idpagina)="" then idpagina = 548	'--- cambiar
	idpagina = EliminaInyeccionSQL(idpagina)
	idpagina = clng(idpagina)

	sql = "SELECT titulo,pagina,numeracion,fecha,fecha_modificacion,tipo,destino FROM WEBISTAS_PAGINAS WHERE idpagina="&idpagina
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	titulo = objRecordset("titulo")	
	pagina = objRecordset("pagina")
	numeracion = objRecordset("numeracion")
	nivel = len(numeracion)
	fechapagina = objrecordset("fecha_modificacion")
	tipopagina = objrecordset("tipo")
	destinopagina = objrecordset("destino")
	seccion = asc(mid(numeracion,3,1))-64
	
	
	'----- Si es restringida y no estás identificado no puedes entrar
	if (tipopagina=7 or tipopagina=8) and session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	
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

	if tipopagina=3 then		'-- Otra página ya existente. Dejo la ruta de esta, pero cambio el título y el contenido
		'response.redirect ("index.asp?idpagina="&destinopagina)
		sql2 = "SELECT titulo,pagina FROM WEBISTAS_PAGINAS WHERE idpagina="&destinopagina
		Set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		set objRecordset2 = OBJConnection.Execute(sql2)
		titulo = objRecordset2("titulo")	
		pagina = objRecordset2("pagina")
	end if

	if isnull(pagina) then saltarpagsig(numeracion)
	if cstr(pagina)="" or cstr(pagina)=" " then saltarpagsig(numeracion)

	'apartado = asc(mid(numeracion,2,1))-64

	'objRecordset.close
	
	titulocompleto = ""
	for i=2 to len(numeracion)
		sql = "SELECT titulo,numeracion,idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='" & mid(numeracion,1,i) & "'"
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		if i<>2 then titulocompleto = titulocompleto & "&nbsp;&gt;&nbsp;" 
		titulocompleto = titulocompleto & "<a href=index.asp?idpagina="&objrecordset("idpagina")&">"&objrecordset("titulo")&"</a>"
	next 
	
	FUNCTION saltarpagsig(codigo)
		sql = "SELECT idpagina FROM WEBISTAS_PAGINAS WHERE numeracion>'"&codigo&"' AND visible=1 ORDER BY numeracion"
		Set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		idpagina = objRecordset("idpagina")
		response.redirect "index.asp?idpagina="&idpagina
	END FUNCTION
	
	FUNCTION vistaprevia(texto)
		if (tipopagina=1 or tipopagina=8) then
			vistaprevia = texto	'-- página de tipo HTML (no aplica códigos)
		else
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
			texto = replace(texto,"<c=","<img src=imagenes/carpetaenlace.gif alt='carpeta de enlaces'>&nbsp;<a class=subtitulo target=_blank href=../../web/vercarpeta.asp?id=")
			texto = replace(texto,"</c>","</a>")
			vistaprevia = texto
		end if
		
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
<title>ISTAS: BBDD Evalúa lo que usas: <%=titulo%></title>
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
<link rel="stylesheet" type="text/css" href="../estructura.css">
</head>
<body>

<div id="contenedor" style="z-index:1;">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja" style="z-index:2;">

			<div id="encabezado_nuevo<% response.write (seccion) %>">
			<table width="100%" cellpadding=0 border=0>
			<tr><td width="215" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="142" height="78" onclick="location.href='index.asp?idpagina=548'" style="cursor:hand">&nbsp;</td>
			    <td width="166" height="78" onclick="location.href='index.asp?idpagina=549'" style="cursor:hand">&nbsp;</td>
			    <td width="160" height="78" onclick="location.href='index.asp?idpagina=550'" style="cursor:hand">&nbsp;</td>
			    <td width="25"  height="78" align="center">
			    	<a href="mailto:ppedroso@istas.ccoo.es?subject=Contacto ECOinformas"><img src="imagenes/ico_contacto.gif" border="0" alt="Contacto"></a><br>
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
              							if cstr(objRecordset("idpagina"))="1018" then response.write "<img src='imagenes/ico_nuevo_cab.gif' alt='nuevo'>&nbsp;"
              							response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&" style='text-decoration:none'>"&lcase(objRecordset("titulo"))&"</a>"
              						end if
              						response.write "</td><td class=textmenusup>|</td>"
							objrecordset.movenext
 						loop %>
              			</tr>
          		</table>
			</div>
			
			<% if request("visita")<>"" then usuario_texto = "Visita virtual" %>
			<% if session("id_ecogente")<>"" and request("visita")="" then
            			sql = "SELECT nombre,apellidos,sexo,confirmado_web,elegible_2007 FROM ECOINFORMAS_GENTE WHERE idgente="&session("id_ecogente")
			   	   				set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	        set objRecordset = OBJConnection.Execute(sql)
		   	   	        if not objRecordset.eof then
		   	   	        	usuario = objRecordset("nombre")&" "&objRecordset("apellidos")
		   	   	        	usuario_sexo = "o"
		   	   	        	if objRecordset("sexo")=75 then usuario_sexo = "a"
		   	   	          usuario_texto = "Usuari" & usuario_sexo & " identificad" & usuario_sexo & ":&nbsp;" & usuario & "&nbsp;"
		   	   	          if isnull(objRecordset("confirmado_web")) then
		   	   	          	confirmado_web = ""
		   	   	          else
		   	   	          	confirmado_web = cstr(objRecordset("confirmado_web"))
		   	   	          end if
		   	   	          if isnull(objRecordset("elegible_2007")) then
		   	   	          	elegible_2007 = ""
		   	   	          else
		   	   	          	elegible_2007 = cstr(objRecordset("elegible_2007"))
		   	   	          end if
		   	   	        end if

      	end if %>

			<div class="textsubmenu" id="submenusup<% response.write (seccion) %>">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
            			<tr><td align="right"><%=usuario_texto%></td></tr>
          		</table>
			</div>
      
			
			<% if len(numeracion)>3 then
			   sql = "SELECT titulo,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE visible=1 AND ((len(numeracion)=5 AND numeracion LIKE '"&mid(numeracion,1,4)&"%')"
			   if len(numeracion)>4 then sql = sql & " OR (len(numeracion)>4 AND numeracion LIKE '"&mid(numeracion,1,5)&"%')"
			   sql = sql & ") ORDER BY numeracion"
			   set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   set objRecordset = OBJConnection.Execute(sql)
		   	   if not objRecordset.eof then
		   	   submenu = 1 %>
		   	   <div id="margen_izquierdo0"><div id="margen_izquierdo<% response.write (seccion) %>">
					<% if idpagina=1165 or idpagina=1179 or idpagina=1180 or idpagina=1181 or idpagina=1182 then %>
					<table cellpadding="2" cellspacing="1" border=0 width="95%" align="center">
					<tr><td class="campo"><img src="imagenes/flecha.gif">&nbsp;<a href="visita_p2007.asp" target="_blank">Visita virtual RISCTOX</a></td></tr>
					</table>
				<% end if %>		   	   	
			<% do while not objRecordset.eof %>
			<table cellpadding="2" cellspacing="1" border=0 width="95%" align="center">
			<tr>
			<% if len(objRecordset("numeracion"))=5 then %>
			<td class="campo"><img src="imagenes/flecha.gif">&nbsp;
			<% else %>
			<td class="campo" width="<%=(len(objRecordset("numeracion"))-5)*12 %>">&nbsp;</td><td class="campo">
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
			   <br>&nbsp;
<% if idpagina=566 then %>
					<form name="busca_enlaces" method="POST" action="buscador_cendoc.asp">
						<table align="center" width="95%" cellpadding=0 cellspacing=5 border=0 style="background-color: #EFEFEF;" class="tabla2">	
							<tr><td>Buscador en el centro de documentación:<br>
								<input type="text" size="18" maxlength="50" class="campo" name="busca_cendoc"><br><input type="submit" class="boton" value="busca">
							</td></tr>
						</table>
					</p>
					</form>			   
<% end if %>			   

			   <p class="texto" style="padding-left: 5px; padding-right: 5px;">Esta página ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundación de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a></p>
			   <% if session("risctox_en_webistas")="si" and idpagina=575 then %>
			   <p class="texto" style="padding-left: 5px; padding-right: 5px;">Esta actividad se realiza  en el marco  del  Convenio de Colaboración suscrito con el Instituto Nacional  de Seguridad  e Higiene en el Trabajo, al amparo de la Resolución de Encomienda  de Gestión de 26 de marzo de 2007, de la Secretaría de Estado de la Seguridad  Social, para el desarrollo de actividades de prevención.</p>
			   <% end if %>
			   <br>&nbsp;
			   </div>
			   
			<% end if
			   end if %>

			<% if submenu=1 or cstr(idpagina)="548" or cstr(idpagina)="965" then %>
			<div id="interiortext">
			<% else %>
			<div id="texto">
			<% end if %>
			
				<div class="texto">
					   				<% if len(numeracion)>3 then
              				     response.write "<br><p class=campo>Est&aacute;s en: "
              				     for i=1 to len(numeracion)-3
              				   	'sql = "SELECT titulo,idpagina FROM WEBISTAS_PAGINAS WHERE visible=1 AND numeracion='"&mid(numeracion,1,2+i)&"'" 
              				   	sql = "SELECT titulo,idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='"&mid(numeracion,1,2+i)&"'" 
			   	   	   	set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	           	set objRecordset = OBJConnection.Execute(sql)
		   	   	           	response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&">"&objRecordset("titulo")&"</a>&nbsp;&gt;&nbsp;"
              				     next
              				     response.write titulo&"</p>"
              				   else
              				     if cstr(idpagina)<>"548" then response.write "<p class=campo>&nbsp;</p>"
              				   end if %>
					
					<% if confirmado_web="2" and cstr(idpagina)="554" and 1=0 then %>
				<table width="50%" class="tabla" cellpadding="10" align="center"><tr><td align="center"><br><a href="formulario_materiales2006.asp">SOLICITAR ENVÍO DE LOS NUEVOS MATERIALES</a><br>&nbsp;</td></tr></table><br><br>
					<% end if %>
					
					<% response.write vistaprevia(pagina)%>
				
				</div>
				<p>&nbsp;</p>
			</div>
			<% if cstr(idpagina)="548" or cstr(idpagina)="965" then %>
			<div id="margen_derecho<% response.write (seccion) %>">
				<center>
				<% if 1=0 then %>
					<div id="identifica2">
						<a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=582"><img src="imagenes/bannercurso06.gif" border=0" alt="Cursos a distancia gratuitos"></a>
					</div>
				<% end if %>
				<% if session("id_ecogente")="" then %>
					<div id="identifica2">
					<form name="form1" id="form1" method="post" action="identifica.asp">
					<input type="hidden" name="idpagina" value="<%=request("idpagina")%>">
					<table width="100%" border="0" cellspacing="2" cellpadding="2" align="center">
				  		<tr bgcolor="#006600">
                					<td>Identificaci&oacute;n</td>
                				</tr>
              					<tr>
                					<td>Clave:</td>
                				</tr>
              					<tr>
                					<td><input name="clave" type="text" class="campoform" id="clave" size="8" maxlenght="20" /></td>
              					</tr>
              					<tr>
                					<td>Contrase&ntilde;a:</td>
                				</tr>
              					<tr>
                					<td><input name="contra" type="password" class="campoform" id="contra" size="8" maxlenght="50" /></td>
              					</tr>
              					<tr>
                					<td><input class="boton" type="submit" name="Submit" value="Enviar" /></td>
              					</tr>
              				</table>
				  	</form>
            				</div>
            				<br><br>
            				<img src=imagenes/entrar.gif align=left><b><a href="formulario.asp?idenc=234" class="campo">Solicitar acceso libre</a></b>
            			<br><br><br><font class="texto"><a href="acceso2.asp">¿olvidaste tu clave y contraseña?</a></font><br><br>
        <% end if %>
        
            			<br><a href="index.asp?idpagina=1165"><img src="imagenes/bocadillovv2.gif" alt="Visita virtual" border="0" style="text-decoration:none"></a><br><br>
									
									<% if cstr(idpagina)="548" or cstr(idpagina)="965" then %>
									<div class="textsubmenu" style="padding:6pt; text-align:left">
									<span class=titulo1>Acceso</span><br>
									Los contenidos de esta página web sirven para toda persona que esté interesada en fomentar un entorno laboral más seguro y sostenible. Para tener acceso a todas la herramientas prácticas que ofrece, puedes darte de alta como usuario/a.<br><br>
									<a href="index.asp?idpagina=560">Saber más sobre ECOinformas</a><br><br>
									<a href="index.asp?idpagina=557">Leer más sobre colectivos específicos a los que va dirigido ECOinformas</a><br><br>
									<% if elegible_2007="2" then %><a href="formulario_materiales2006.asp"><img src="imagenes/bocadillomateriales.gif" alt="Solicita los nuevos materiales" border="0" style="text-decoration:none"></a><% end if %>
									<a href="index.asp?idpagina=554"><img src="imagenes/bocadillomateriales2.gif" alt="Descarga los nuevos materiales" border="0" style="text-decoration:none"></a>
									</div>
									
            			<% end if %>
            					
				
				</center>
			</div>				
			<% end if %>
			
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
<% if idpagina<>1113 then %>
<script language="JScript" type="text/jscript" src="activateActiveX.js"></script>
<% end if %>
</body>
</html>