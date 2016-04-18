<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

	
	
	'----- Si es restringida y no estás identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina=576"
	'---- ATENCIÓN: ponerlo cuando publiquemos en abierto
	
	'numeracion = "AICDA"
	idpagina = 633	'--- página Sectores (sólo para registrar estadísticas)
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
              				     response.write "Sectores</p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              			   end if %>
				
				<p class=titulo3>Base de datos de alternativas de sustitución de productos con riesgo tóxico</p>

				

<%				registrosporpagina = 25
				if request("pag")<>"" then
			   		Session("pagina") = request("pag")
				else
   					Session("pagina") = 1
				end if

				if request("ordenacion")="" then 
					ordenacion = "nombre"
				else
					ordenacion = request("ordenacion")
				end if
				if right(ordenacion,4)<>"DESC" then texto_ord = " DESC"
					
				sql = "SELECT * FROM RQ_SECTORES ORDER BY sector"
				set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset.Open sql,objConnection,adOpenKeyset
		   	   	registros = objRecordset.recordCount
	   	   		objrecordset.PageSize = registrosporpagina
		   	   	if registros=0 then 
		   	   		response.write "<table class=tabla3 width='90%' align=center border=0 cellpadding=4 cellspacing=0>"
		   	   		response.write "<tr><td class=texto colspan=4>No hay procesos</td></tr>"
		   	   	else 
		   	   		response.write "<p class=texto align=center>Pulsa sobre el nombre del sector para ver las alternativas relacionadas:</p>"
		   	   		response.write "<table class=tabla3 width='90%' align=center border=0 cellpadding=4 cellspacing=0>"
		   	   	%>
				  <tr>
				  	<td class="subtitulo3" colspan=4>
				  		<table width="100%" align="center"><tr><td>Sectores (<%=registros%>)&nbsp;</td>
				  		<td align=right><img src="imagenes/ico_alt_sectores.gif"></td></tr></table>
				  	</td>
				  </tr>
				  <tr>
					<td class="celda_risctox" align="right"><b>CNAE</b></td>
					<td class="celda_risctox"><b>Nombre del sector</b></td>
				  </tr>				  
<%				objRecordset.movefirst
				objrecordset.AbsolutePage = Session("pagina")
				i = 0
				do while not objRecordset.eof and i<registrosporpagina %>
				<tr>
					<td class="celda_risctox" align="right"><a href="alternativas.asp?sector=<%=objRecordset("idsector")%>&nombre=<%=objRecordset("sector")%>"><%=objRecordset("codigo")%></a>&nbsp;</td>
					<td class="celda_risctox"><a href="alternativas.asp?sector=<%=objRecordset("idsector")%>&nombre=<%=objRecordset("sector")%>" title="<%=objRecordset("sector")%>"><%=formato(objRecordset("sector"),300)%></a>&nbsp;</td>
				</tr>
<%				objRecordset.movenext
				i = i+1
				loop
				end if %>							
				<tr><td class=celda_risctox colspan=4>&nbsp;</td></tr>
<% 				if objRecordset.Pagecount>1 then%>
				<tr><td class=texto colspan=4 align="center">Hay <%=registros%> sectores que cumplen los criterios de búsqueda. Se muestran sólo <%=registrosporpagina%> por página.</td></tr>
				<tr><td class=texto colspan=4 align="center">
<%				if Clng(Session("pagina")) > 1 then %>
				<a href="sectores.asp?ordenacion=<%=ordenacion%>&buscar=<%=buscar%>&pag=<%=Clng(Session("pagina"))-1%>">anterior&nbsp;&lt;&lt;</a>&nbsp;&nbsp;&nbsp;
<%				end if %>
				Página&nbsp;<%=Session("pagina")%>&nbsp;de&nbsp;<%=objRecordset.Pagecount%>
<%				if Clng(Session("pagina")) < objRecordset.Pagecount then %>
				&nbsp;&nbsp;&nbsp;<a href="sectores.asp?ordenacion=<%=ordenacion%>&buscar=<%=buscar%>&pag=<%=Clng(Session("pagina"))+1%>">&gt;&gt;&nbsp;siguiente</a>
<%				end if %>
				<form name="form_pag">Cambiar a la página:&nbsp;<input type="text" size="2" maxlenght="2" class="campo" value="<%=Session("pagina")%>" name="pag_ir">&nbsp;<input type="button" class="boton" value="ir" onclick="cambia_pag()"></form>
<%				end if %>				
				</td></tr>				
				</table>

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
<script>
function cambia_pag()
{	if (form_pag.pag_ir.value><%=objRecordset.Pagecount%>)
	{ alert('Sólo se puede ir hasta la página <%=objRecordset.Pagecount%>'); }
	else
	{ location.href='sectores.asp?ordenacion=<%=ordenacion%>&buscar=<%=buscar%>&pag='+form_pag.pag_ir.value; }
}
</script>