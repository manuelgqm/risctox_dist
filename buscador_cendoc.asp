<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	
	'----- Si es restringida y no est�s identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	
	numeracion = "AIBCE"
	seccion = asc(mid(numeracion,3,1))-64

	idpagina = 1178	'--- p�gina buscador del Centro de dcumentaci�n (s�lo para registrar estad�sticas)
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

	sql = "SELECT numeracion FROM ENL_TEMAS WHERE id=332"	'-- Materiales del proyecto
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	numeracion1 = objrecordset("numeracion")

	sql = "SELECT numeracion FROM ENL_TEMAS WHERE id=329"	'-- Actividades de formaci�n
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	numeracion2 = objrecordset("numeracion")

	sql = "SELECT numeracion FROM ENL_TEMAS WHERE id=341"	'-- Servicio de documentaci�n (Vig.Tecnol�gica y Enlaces web)
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	numeracion3 = objrecordset("numeracion")

	sql = "SELECT numeracion FROM ENL_TEMAS WHERE id=716"	'--Enlaces Observatorio ambiental
	set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	numeracion4 = objrecordset("numeracion")
	
	titulocompleto = ""
	for i=2 to len(numeracion)
		sql = "SELECT titulo,numeracion,idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='" & mid(numeracion,1,i) & "'"
		set objRecordset = Server.CreateObject ("ADODB.Recordset")
		set objRecordset = OBJConnection.Execute(sql)
		titulo = objrecordset("titulo")
		if i<>2 then titulocompleto = titulocompleto & "&nbsp;&gt;&nbsp;" 
		if not objrecordset.eof then titulocompleto = titulocompleto & "<a href=index.asp?idpagina="&objrecordset("idpagina")&">"&objrecordset("titulo")&"</a>"
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
<title>ECOinformas: <%=titulo%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multim�dia" />
<meta name="description" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />
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
			
			
			<% if len(numeracion)>3 then
			   sql = "SELECT titulo,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE visible=1 AND ((len(numeracion)=5 AND numeracion LIKE '"&mid(numeracion,1,4)&"%')"
			   if len(numeracion)>4 then sql = sql & " OR (len(numeracion)>4 AND numeracion LIKE '"&mid(numeracion,1,5)&"%')"
			   sql = sql & ") ORDER BY numeracion"
			   set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   set objRecordset = OBJConnection.Execute(sql)
		   	   if not objRecordset.eof then
		   	   submenu = 1 %>
		   	   <div id="margen_izquierdo0"><div id="margen_izquierdo<% response.write (seccion) %>">
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
<% if idpagina=1178 then %>
					<form name="busca_enlaces" method="POST" action="buscador_cendoc.asp">
						<table align="center" width="95%" cellpadding=0 cellspacing=5 border=0 style="background-color: #EFEFEF;" class="tabla2">	
							<tr><td>Buscador en el centro de documentaci�n:<br>
								<input type="text" size="18" maxlength="50" class="campo" name="busca_cendoc" value="<%=request("busca_cendoc")%>"><br><input type="submit" class="boton" value="busca">
							</td></tr>
						</table>
					</p>
					</form>			   
<% end if %>			   

			   <p class="texto" style="padding-left: 5px; padding-right: 5px;">Esta p�gina ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundaci�n de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a></p>
			   <br>&nbsp;
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
		   	   	           	if not objrecordset.eof then response.write "<a href=index.asp?idpagina="&objRecordset("idpagina")&">"&objRecordset("titulo")&"</a>&nbsp;&gt;&nbsp;"
              				     next
              				     response.write "Resultado de la b�squeda</p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              				   end if
              				   objRecordset.close %>
		<p class="titulo2">Buscando:&nbsp;<%=request("busca_cendoc")%></p>
						
<!-- buscador de p�ginas-->				
				<% 
			   		   sql = "SELECT titulo,idpagina,numeracion,tipo FROM WEBISTAS_PAGINAS WHERE numeracion LIKE '"&mid(numeracion,1,4)&"%' AND (tipo=2 or tipo=7) AND visible=1"
			   		   sql = sql & " AND (titulo LIKE '%"&request("busca_cendoc")&"%' OR pagina LIKE '%"&request("busca_cendoc")&"%') "
			   		   sql = sql & " ORDER BY numeracion"
					   	objRecordset.Open sql,objConnection,adOpenKeyset
			   		   num_resultados = objRecordset.recordcount %>
<p class=texto"><b>Hay <%=num_resultados%> p�ginas.<% if num_resultados>0 then %> Pulsa sobre una de ellas para visitarla<%else%> Prueba con otra palabra<% end if %></b></p>
							   	   		   
<%		   	   		   if not objRecordset.eof then
		   	   		     r = 0
		   	   		     do while not objRecordset.eof
		   	   		     numeracion = objRecordset("numeracion")
		   	   		     r = r+1 %>
						<p class="texto"><%=r%>.&nbsp;
						<% if len(numeracion)>3 then
              				     		for i=1 to len(numeracion)-3
              				   			sql2 = "SELECT titulo,idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='"&mid(numeracion,1,2+i)&"'" 
			   	   	   			set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	           			set objRecordset2 = OBJConnection.Execute(sql2)
		   	   	           			response.write objRecordset2("titulo")&"&nbsp;&gt;&nbsp;"
              				     		next
              				   	   end if %>
						<a href="index.asp?idpagina=<%=objRecordset("idpagina")%>"><%=objRecordset("titulo")%></a>
						<% if objRecordset("tipo")=7 then response.write "&nbsp;<img src='imagenes\candado.gif' border=0 alt='P�gina restringida'>" %></p>
<%   					     objRecordset.movenext
			   		     loop %>
						<p align=center class="tabla2"><img src='imagenes/candado.gif' border=0 alt='P�gina restringida'>&nbsp;P�ginas restringidas. Se requiere clave para acceder que se puede solicitar pulsando <a href="formulario2.asp">aqu�</a></p>			   		     
			   		 <%  else %>
			   		   <p class="texto">No hay ninguna p�gina coincidente con esta b�squeda</p>
  					 <p>&nbsp;</p>
					   <p>&nbsp;</p>

			   		<% end if %>
			   	 
					
<!--/buscador de p�ginas-->					

<!--/buscador de enlaces-->
					<% sql = "SELECT DISTINCT ENL_ENLACES.Id,ENL_ENLACES.afiliacion,ENL_ENLACES.titulo,ENL_ENLACES.url FROM ENL_ENLACES LEFT JOIN ENL_CLASIFICACION ON ENL_ENLACES.Id = ENL_CLASIFICACION.enlace LEFT JOIN ENL_TEMAS ON ENL_CLASIFICACION.tema = ENL_TEMAS.id WHERE (ENL_TEMAS.numeracion LIKE '"&numeracion1&"%' OR ENL_TEMAS.numeracion LIKE '"&numeracion2&"%' OR ENL_TEMAS.numeracion LIKE '"&numeracion3&"%' OR ENL_TEMAS.numeracion LIKE '"&numeracion4&"%') AND (ENL_ENLACES.afiliacion LIKE '%"&request("busca_cendoc")&"%' OR ENL_ENLACES.titulo LIKE '%"&request("busca_cendoc")&"%' OR ENL_ENLACES.subtitulo LIKE '%"&request("busca_cendoc")&"%' OR ENL_ENLACES.indice LIKE '%"&request("busca_cendoc")&"%') ORDER BY ENL_ENLACES.id DESC"
			   		   'response.write "<!--"&sql&"-->"
			   		   set objRecordset = Server.CreateObject ("ADODB.Recordset")
			   		   objRecordset.Open sql,objConnection,adOpenKeyset
			   		   num_resultados = objRecordset.recordcount %>

					<p class=texto"><b>Hay <%=num_resultados%> enlaces.<% if num_resultados>0 then %> Pulsa sobre uno de ellos para visitarlo<%else%> Prueba con otra palabra<% end if %></b></p>
					<table align="center" cellpadding="3" cellspacing="3">
			   		<% i=1
			   		   do while not objRecordset.eof
			   		   if i=1 then 
			   		   	color = "DDDDDD"
			   		   else
			   		   	color = "F0F0F0"
			   		   end if %>
					<tr><td class="texto" valign="top" align="right"><a href="abreenlacer.asp?idenlace=<%=objRecordset("Id")%>" target="_blank"><img src="imagenes/ico_puntito.gif" valign="top" border=0></a></td><td class="texto" bgcolor="#<%=color%>"><b><%=objRecordset("afiliacion")%></b><br><a href="abreenlacer.asp?idenlace=<%=objRecordset("Id")%>" target="_blank"><%=mid(objRecordset("titulo"),1,40)%><br><a href="abreenlacer.asp?idenlace=<%=objRecordset("Id")%>" target="_blank"><%=mid(objRecordset("url"),1,40)%></td></tr>
					<% i=-1*i
					   objRecordset.movenext
					   loop %>
					</table>
					<p align="center"><input type="button" class="boton" value="imprimir" onclick="print()"></p>



				
				</div>
				<p>&nbsp;</p>
			</div>
			
			<map name="Map1" id="Map1">
            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
            		<area shape="rect" coords="400,38,546,104" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="547,38,701,104" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<map name="Map2" id="Map2">
            		<area shape="rect" coords="300,8,392,66" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
            		<area shape="rect" coords="393,8,539,66" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,8,694,66" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
      			<map name="Map3" id="Map3">
            		<area shape="rect" coords="300,18,392,80" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
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