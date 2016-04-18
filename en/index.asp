<%
'if instr(request.ServerVariables("http_host"),"risctox.istas.net")<>0 then response.redirect "http://www.istas.net/risctox"
%>
<!--#include file="../EliminaInyeccionSQL.asp"-->
<!--#include file="../dn_conexion.asp"-->

<%

	if (request.servervariables("SERVER_NAME") = "217.13.81.22") then
		if (request("idpagina")<>"") then
			response.redirect("http://www.istas.net/risctox/index.asp?idpagina="&request("idpagina"))
		else
			response.redirect("http://www.istas.net/risctox/index.asp")
		end if
	end if


	idpagina = request("idpagina")

	if cstr(idpagina)="" then idpagina = 575	'--- cambiar
	idpagina = EliminaInyeccionSQL(idpagina)
	idpagina = clng(idpagina)

	sql = "SELECT titulo,pagina,titulo_eng,pagina_eng,numeracion,fecha,fecha_modificacion,tipo,destino FROM WEBISTAS_PAGINAS WHERE idpagina="&idpagina
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	titulo = objRecordset("titulo_eng")
	pagina = objRecordset("pagina_eng")
	numeracion = objRecordset("numeracion")
	nivel = len(numeracion)
	fechapagina = objrecordset("fecha_modificacion")
	tipopagina = objrecordset("tipo")
	destinopagina = objrecordset("destino")
	seccion = asc(mid(numeracion,3,1))-64



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
			texto = replace(texto,"ecoinformas/web","ecoinformas08/web")
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
<title>ISTAS: <%=titulo%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="Risctox" />
<meta name="Author" content="SPL Sistemas de Información - www.spl-ssi.com" />
<meta name="description" content="" />
<meta name="Subject" content="" />
<meta name="Keywords" content="" />
<meta name="Language" content="English" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
<link rel="stylesheet" type="text/css" href="../estructura.css">
<link rel="stylesheet" type="text/css" href="css/en.css">
</head>
<body>

<div id="contenedor" style="z-index:1;">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja" style="z-index:2;">


<!--#include file="dn_cabecera.asp"-->

			<%
			   if len(numeracion)>3 then
			   sql = "SELECT titulo_eng,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE visible=1 AND titulo_eng<>'' AND ((len(numeracion)=5 AND numeracion LIKE '"&mid(numeracion,1,4)&"%')"
			   if len(numeracion)>4 then sql = sql & " OR (len(numeracion)>4 AND numeracion LIKE '"&mid(numeracion,1,5)&"%')"
			   sql = sql & ") ORDER BY numeracion"
			   set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   set objRecordset = OBJConnection.Execute(sql)
		   	   if not objRecordset.eof then
		   	   submenu = 1
		   	%>
		   	   <div id="margen_izquierdo0">
		   	   	<div id="margen_izquierdo<% response.write (seccion) %>">

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
				response.write "<b>"&objRecordset("titulo_eng")&"</b>"
			   else
			   	response.write objRecordset("titulo_eng")
			   end if %>
			</a>
            <%
			 'Sergio: Añado ototóxicos detrás de neurotóxicos
			 'SPL: No sé porqué está este documento introducido a mano Ototóxicos y Alérgenos, pero lo elimino por indicación de Elena y Tatiana (10/10/2012)
			 if objRecordset("idpagina") = 6119999999999999 then %>
			 	<br />&nbsp;&nbsp;&nbsp;&nbsp;<img src="imagenes/flecha.gif">&nbsp; <a href="index.asp?idpagina=<%=objRecordset("idpagina")%>">Ototóxicos</a>

			<%
			end if

			%>
            <%
			 'Sergio: Añado alérgentos REACH detrás de sensibilizantes
			 'SPL: No sé porqué está este documento introducido a mano Ototóxicos y Alérgenos, pero lo elimino por indicación de Elena y Tatiana (10/10/2012)
			 if objRecordset("idpagina") = 6129999999999 then %>
			 	<br />&nbsp;&nbsp;&nbsp;&nbsp;<img src="imagenes/flecha.gif">&nbsp; <a href="index.asp?idpagina=<%=objRecordset("idpagina")%>">Alérgenos REACH</a>

			<%
			end if

			%>
			</td></tr></table>
			<%
			   objRecordset.movenext
			   loop
			%>
			</div>
			   <br>&nbsp;


			   <p class="texto" style="padding-left: 5px; padding-right: 5px;">This site has been developed by <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> - <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a></p>
			   <% if 1=1 and idpagina=575 then %>
			   <p class="texto" style="padding-left: 5px; padding-right: 5px;">This activity has been commisioned by the <a href="http://www.etui.org/" target="_blank"><b>European Trade Union Institute (ETUI)</b></a> and supported by the <a target="_blank" href="http://www.eeb.org/" style="font-weight:bold">European Environmental Bureau (EEB)</a></p>
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

					   				<% if 1=0 and len(numeracion)>3 then
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

			<img src="imagenes/pie_risctox.gif" width="708" border="0">


    		</div>
    	</div>
	<div id="sombra_abajo" ><p class="texto" style="padding-left: 5px; padding-right: 5px;color:#999;">This web has been developed by <a href="http://www.spl-ssi.com" style="color:#999;" target="_blank">SPL Sistemas de Informaci&oacute;n</a></p></div>
</div>
<% if idpagina<>1113 then %>
<script language="JScript" type="text/jscript" src="../activateActiveX.js"></script>
<% end if %>
<!--#include file="../../cookie_accept_en.asp" -->
</body>
</html>