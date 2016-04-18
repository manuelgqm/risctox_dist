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
	
	'numeracion = "AICDA"
	idpagina = 635	'--- página Alternativa (sólo para registrar estadísticas)
	sql = "SELECT numeracion FROM WEBISTAS_PAGINAS WHERE idpagina="&idpagina
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	numeracion = objRecordset("numeracion")
	
	
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
			if len(x)>(lon-3) then x = mid(x,1,lon-3)&"..."
			formato = x
		end if
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
<% if request("paso")<>"" then %>
<div style="position:absolute; width=778px; height=400px; top:0px; z-index:3; align=center">
<table width="778" align="center"><tr><td align="center">
<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" width="778" height="714" id="paso<%=request("paso")%>" align="middle">
<param name="allowScriptAccess" value="sameDomain" />
<param name="movie" value="paso<%=request("paso")%>.swf">
<param name="quality" value="high" />
<param name="wmode" value="transparent" />
<param name="loop" value="false" />
<param name="menu" value="false" />
<embed src="paso<%=request("paso")%>.swf" width="778" height="714" quality="high" wmode="transparent" bgcolor="#ffffff" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
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
			

<%	id = request("id")
	if id="" then id=0
	dim campo(8,2)
	campo(1,0)="idalternativa"        
	campo(2,0)="alternativa"        
	campo(3,0)="comentarios"    
	campo(4,0)="url1"    
	campo(5,0)="url2"    
	campo(6,0)="url3"    
	campo(7,0)="url4"    
	campo(8,0)="url5"    
	

	if id<>"" and cstr(id)<>"0" then
	orden = "SELECT * FROM RQ_ALTERNATIVAS WHERE idalternativa="&clng(id)
	Set dorga = Server.CreateObject ("ADODB.Recordset")
	dorga.Open orden,objConnection,adOpenKeyset

	for i=1 to 8
		campo(i,1) = dorga(campo(i,0))
		'response.write campo(i,0)&"= "&campo(i,1)&"<br>"
	next
	end if

%>

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
              				     response.write "<a href=alternativas.asp>Listado</a>&nbsp;&gt;&nbsp;Alternativa:&nbsp;"&ucase(campo(2,1))&"</p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              			   end if %>
				
				<p class=titulo3>Base de datos de alternativas de sustitución de productos con riesgo tóxico</p>

<table width="100%" class="tabla3">
    <tr>
      <td class="titulo3" align="right">Alternativa:</td>
      <td class="texto"><b><%=campo(2,1)%></b></td>
    </tr>
    <tr>
      <td class="titulo3" align="right">Comentarios:</td>
      <td class="texto"><%=campo(3,1)%></td>
    </tr>
<% for i=1 to 5%>
<% if campo(3+i,1)<>"" then
	enlace = campo(3+i,1)
	enlace = replace(enlace,"http://www.istas.net/RQ/ficheros/","")
	enlace = replace(enlace,"%20"," ")
 %>
    <tr>
      <td class="titulo3" align="right">Enlace <%=i%>:</td>
      <td class="texto"><a href="<%=campo(3+i,1)%>" target="_blank"><%=enlace%></a></td>
    </tr> 
<% end if %>   
<% next %>    
    
</table>
<br>&nbsp;

<%'-- Sectores relacionados ------------------------ %>
<% 	orden = "SELECT RQ_SECTORES.idsector,RQ_SECTORES.sector,RQ_SECTORES.codigo FROM RQ_SECTORES LEFT JOIN RQ_ALTERNATIVAS_RELACIONES ON RQ_SECTORES.idsector=RQ_ALTERNATIVAS_RELACIONES.id_relacion WHERE RQ_ALTERNATIVAS_RELACIONES.tabla_relacion='RQ_SECTORES' AND RQ_ALTERNATIVAS_RELACIONES.idalternativa="&id
	Set dorga = Server.CreateObject ("ADODB.Recordset")
	dorga.Open orden,objConnection,adOpenKeyset
	cuantos_sectores = dorga.recordcount

	if cuantos_sectores>0 then
%>	
<table width="100%" class="tabla3">
<tr>
  <td class="titulo3" colspan=2>Sectores relacionados:&nbsp;<%=cuantos_sectores%></td>
</tr>
<%	do while not dorga.eof %>
<tr>
<td class="texto" align="left" width="20%" valign="top">CNAE:&nbsp;<%=dorga("codigo")%>&nbsp;</td>
<td class="texto"><%=mid(dorga("sector"),1,300)%></td>
</tr>
<%	dorga.movenext
	loop %>
</table>
<br>&nbsp;
<%	end if %>

<%'-- Procesos relacionados ------------------------ %>
<% 	orden = "SELECT RQ_PROCESOS.proceso,RQ_PROCESOS.idproceso FROM RQ_PROCESOS LEFT JOIN RQ_ALTERNATIVAS_RELACIONES ON RQ_PROCESOS.idproceso=RQ_ALTERNATIVAS_RELACIONES.id_relacion WHERE RQ_ALTERNATIVAS_RELACIONES.tabla_relacion='RQ_PROCESOS' AND RQ_ALTERNATIVAS_RELACIONES.idalternativa="&id
	Set dorga = Server.CreateObject ("ADODB.Recordset")
	dorga.Open orden,objConnection,adOpenKeyset
	cuantos_procesos = dorga.recordcount

	if cuantos_procesos>0 then
%>	
<table width="100%" class="tabla3">
<tr>
  <td class="titulo3" colspan=2>Procesos relacionados:&nbsp;<%=cuantos_procesos%></td>
</tr>
<%	do while not dorga.eof %>
<tr>
<td class="texto" align="right" width="20%">&nbsp;</td>
<td class="texto"><%=mid(dorga("proceso"),1,300)%></td>
</tr>
<%	dorga.movenext
	loop %>
</table>
<br>&nbsp;
<%	end if %>

<%'-- Produtos relacionados ------------------------ %>
<% 	orden = "SELECT RQ_PRODUCTOS.producto,RQ_PRODUCTOS.idproducto FROM RQ_PRODUCTOS LEFT JOIN RQ_ALTERNATIVAS_RELACIONES ON RQ_PRODUCTOS.idproducto=RQ_ALTERNATIVAS_RELACIONES.id_relacion WHERE RQ_ALTERNATIVAS_RELACIONES.tabla_relacion='RQ_PRODUCTOS' AND RQ_ALTERNATIVAS_RELACIONES.idalternativa="&id
	Set dorga = Server.CreateObject ("ADODB.Recordset")
	dorga.Open orden,objConnection,adOpenKeyset
	cuantos_productos = dorga.recordcount

	if cuantos_productos>0 then
%>	
<table width="100%" class="tabla3">
<tr>
  <td class="titulo3" colspan=2>Productos relacionados:&nbsp;<%=cuantos_productos%></td>
</tr>
<%	do while not dorga.eof %>
<tr>
<td class="texto" align="right" width="20%">&nbsp;</td>
<td class="texto"><%=mid(dorga("producto"),1,300)%></td>
</tr>
<%	dorga.movenext
	loop %>
</table>
<br>&nbsp;
<%	end if %>

<%'-- Listas relacionadas ------------------------ %>
<% 	if id_ecogente=179 then 
 		orden = "SELECT RQ_LISTAS.lista,RQ_LISTAS.url,RQ_LISTAS.idlista FROM RQ_LISTAS LEFT JOIN RQ_ALTERNATIVAS_RELACIONES ON RQ_LISTAS.idlista=RQ_ALTERNATIVAS_RELACIONES.id_relacion WHERE RQ_ALTERNATIVAS_RELACIONES.tabla_relacion='RQ_LISTAS' AND RQ_ALTERNATIVAS_RELACIONES.idalternativa="&id
   	else
   		orden = "SELECT RQ_LISTAS.lista,RQ_LISTAS.url,RQ_LISTAS.idlista FROM RQ_LISTAS LEFT JOIN RQ_ALTERNATIVAS_RELACIONES ON RQ_LISTAS.idlista=RQ_ALTERNATIVAS_RELACIONES.id_relacion WHERE RQ_ALTERNATIVAS_RELACIONES.tabla_relacion='RQ_LISTAS' AND RQ_ALTERNATIVAS_RELACIONES.idalternativa="&id&" AND RQ_LISTAS.visible=1"
   	end if
	Set dorga = Server.CreateObject ("ADODB.Recordset")
	dorga.Open orden,objConnection,adOpenKeyset
	cuantas_listas = dorga.recordcount

	if cuantas_listas>0 then
%>	
<table width="100%" class="tabla3">
<tr>
  <td class="titulo3" colspan=2>Listas relacionadas:&nbsp;<%=cuantas_listas%></td>
</tr>
<%	do while not dorga.eof %>
<tr>
<td class="texto" align="right" width="20%">&nbsp;</td>
<td class="texto"><a href="<%=dorga("url")%>"><%=dorga("lista")%></a></td>
</tr>
<%	dorga.movenext
	loop %>
</table>
<br>&nbsp;
<%	end if %>




<%'-- Sustancias relacionadas ------------------------ %>
<% 	orden = "SELECT RQ_SUSTANCIAS.nombre,RQ_SUSTANCIAS.CAS,RQ_SUSTANCIAS.id FROM RQ_SUSTANCIAS LEFT JOIN RQ_ALTERNATIVAS_RELACIONES ON RQ_SUSTANCIAS.id=RQ_ALTERNATIVAS_RELACIONES.id_relacion WHERE RQ_ALTERNATIVAS_RELACIONES.tabla_relacion='RQ_SUSTANCIAS' AND RQ_ALTERNATIVAS_RELACIONES.idalternativa="&id
	Set dorga = Server.CreateObject ("ADODB.Recordset")
	dorga.Open orden,objConnection,adOpenKeyset
	cuantos_sustancias = dorga.recordcount

	if cuantos_sustancias>0 then
%>	
<table width="100%" class="tabla3">
<tr>
  <td class="titulo3" colspan="2">Sustancias relacionadas:&nbsp;<%=cuantos_sustancias%></td>
</tr>
<%	do while not dorga.eof %>
<tr>
<td class="texto" align="right" width="20%"><a href="risctox3.asp?cas=<%=dorga("cas") %>"><% if instr(dorga("cas"),"-XX-")=0 then response.write "["&dorga("cas")&"]"%></a></td><td class="negro"><a href="risctox3.asp?cas=<%=dorga("cas") %>"><%=mid(dorga("nombre"),1,300)%></a></td>
</tr>
<%	dorga.movenext
	loop %>
</table>
<br>&nbsp;
<% 	end if %>

				<% if 1=0 then %><p align="center"><input type="button" class="boton" value="imprimir listado completo" onclick="print()"></p><% end if %>
				</div>
				</div>
				<p align=center><object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" width=530 height=100 id="anima0" align="middle"><param name="allowScriptAccess" value="sameDomain" /><param name="movie" value="http://www.istas.net/recursos/ANI/ISTAS_01079.swf" />
				<param name="quality" value="high" /><param name="wmode" value="transparent" /><param name="bgcolor" value="#ffffff" /><embed src="http://www.istas.net/recursos/ANI/ISTAS_01079.swf" quality="high" wmode="transparent" bgcolor="#ffffff" width=530 height=100 name="" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" /></object></p>
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
