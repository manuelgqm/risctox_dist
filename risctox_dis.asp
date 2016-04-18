<% response.redirect "dn_risctox_dis.asp" %>
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

	
	
	'----- Si es restringida y no estás identificado no puedes entrar
	if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
	'---- ATENCIÓN: ponerlo cuando publiquemos en abierto
	
	numeracion = "AICCA"
	
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
              				     response.write "<a href=risctox1.asp>Inicio</a>&nbsp;&gt;&nbsp;Disruptores endocrinos</p>"
              				   else
              				     response.write "<p class=campo>&nbsp;</p>"
              			   end if %>
				
				<p class=titulo3>RISCTOX: Disruptores endocrinos</p>

				<form name="form_buscar" action="risctox_dis.asp">
				<table class="tabla3" width="90%" align="center">
				  <tr>
					<td class="subtitulo3" nowrap align="right">Buscador de sustancias disruptores endocrinos:&nbsp;</td>
					<td class="texto"><input type="text" size="30" class="campo" name="buscar" value="<%=request("buscar")%>">&nbsp;<input type="button" class="boton" value="buscar" onclick="enviar()"></td>
				  </tr>
				  <tr>
					<td class="texto">&nbsp;</td>
					<td class="texto">(puedes escribir una parte del nombre o sinónimo, o el CAS, CE o RD)</td>
				  </tr>
				</table>
				</form>
				

<%				buscar = request("buscar")
				buscar = unquote(buscar)
				registrosporpagina = 25
				if request("pag")<>"" then
			   		Session("pagina") = request("pag")
				else
   					Session("pagina") = 1
				end if

				if request("ordenacion")="" then 
					ordenacion = "RQ_SUSTANCIAS_DIS.nombre"
				else
					ordenacion = request("ordenacion")
				end if
				if right(ordenacion,4)<>"DESC" then texto_ord = " DESC"
					
				'sql = "SELECT RQ_SUSTANCIAS.*,RQ_SUSTANCIAS_DIS.fuente FROM RQ_SUSTANCIAS RIGHT JOIN RQ_SUSTANCIAS_DIS ON RQ_SUSTANCIAS.CAS=RQ_SUSTANCIAS_DIS.CAS "
				'sql = sql & "WHERE RQ_SUSTANCIAS.nombre<>'' AND (nombre LIKE '%"&buscar&"%' OR sinonimos LIKE '%"&buscar&"%' OR RQ_SUSTANCIAS_DIS.cas LIKE '%"&buscar&"%' OR cee LIKE '%"&buscar&"%' OR rd LIKE '%"&buscar&"%') ORDER BY "&ordenacion
				
				sql = "SELECT RQ_SUSTANCIAS.NOMBRE, RQ_SUSTANCIAS_DIS.CAS, RQ_SUSTANCIAS_DIS.nombre as SUSTANCIA, RQ_SUSTANCIAS_DIS.fuente FROM RQ_SUSTANCIAS RIGHT JOIN RQ_SUSTANCIAS_DIS ON RQ_SUSTANCIAS.CAS=RQ_SUSTANCIAS_DIS.CAS "
				if buscar<>"" then sql = sql & "WHERE (RQ_SUSTANCIAS_DIS.CAS LIKE '%"&buscar&"%' OR RQ_SUSTANCIAS_DIS.nombre LIKE '%"&buscar&"%' OR RQ_SUSTANCIAS.nombre LIKE '%"&buscar&"%' OR RQ_SUSTANCIAS.sinonimos LIKE '%"&buscar&"%' OR RQ_SUSTANCIAS.cas LIKE '%"&buscar&"%' OR RQ_SUSTANCIAS.cee LIKE '%"&buscar&"%' OR RQ_SUSTANCIAS.rd LIKE '%"&buscar&"%') "
				sql = sql & "ORDER BY "&ordenacion

				
				set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   	objRecordset.Open sql,objConnection,adOpenKeyset
		   	   	registros = objRecordset.recordCount
	   	   		objrecordset.PageSize = registrosporpagina
		   	   	if registros=0 then 
		   	   		response.write "<table class=tabla3 width='90%' align=center border=0 cellpadding=4 cellspacing=0>"
		   	   		response.write "<tr><td class=texto colspan=4>No hay sustancias con este criterio de búsqueda</td></tr>"
		   	   	else 
		   	   		response.write "<p class=texto align=center>Pulsa sobre el nombre de la sustancia para ver su ficha completa<br>y sobre los títulos de columna para reordenar el listado:</p>"
		   	   		response.write "<table class=tabla3 width='90%' align=center border=0 cellpadding=4 cellspacing=0>"
		   	   	%>
				  <tr>
				  	<td class="subtitulo3">
				  		<a href="risctox_dis.asp?buscar=<%=buscar%>&ordenacion=RQ_SUSTANCIAS_DIS.nombre<%=texto_ord%>">Sustancias (<%=registros%>)</a>&nbsp;
				  		<% if ordenacion="RQ_SUSTANCIAS_DIS.nombre DESC" then response.write "<img src=imagenes/abajo.gif>" %>
				  		<% if ordenacion="RQ_SUSTANCIAS_DIS.nombre" then response.write "<img src=imagenes/arriba.gif>" %>
				  	</td>
				  	<td class="subtitulo3">
				  		<a href="risctox_dis.asp?buscar=<%=buscar%>&ordenacion=RQ_SUSTANCIAS_DIS.cas<%=texto_ord%>">CAS</a>&nbsp;
				  		<% if ordenacion="RQ_SUSTANCIAS_DIS.cas DESC" then response.write "<img src=imagenes/abajo.gif>" %>
				  		<% if ordenacion="RQ_SUSTANCIAS_DIS.cas" then response.write "<img src=imagenes/arriba.gif>" %>
				  	</td>
				  	<td class="subtitulo3">
				  		<a href="risctox_dis.asp?buscar=<%=buscar%>&ordenacion=RQ_SUSTANCIAS_DIS.fuente<%=texto_ord%>">FUENTE</a>&nbsp;
				  		<% if ordenacion="RQ_SUSTANCIAS_DIS.fuente DESC" then response.write "<img src=imagenes/abajo.gif>" %>
				  		<% if ordenacion="RQ_SUSTANCIAS_DIS.fuente" then response.write "<img src=imagenes/arriba.gif>" %>
				  	</td>
				  </tr>
<%				objRecordset.movefirst
				objrecordset.AbsolutePage = Session("pagina")
				i = 0
				do while not objRecordset.eof and i<registrosporpagina
					cas_actual = objRecordset("cas")
					if objRecordset("nombre")<>"" then
						nombre_actual = objRecordset("nombre")
					else
						nombre_actual = objRecordset("sustancia")
					end if %>
				<tr>
					<td class="celda_risctox"><a href="risctox3.asp?buscar=<%=buscar%>&cas=<%=cas_actual%>&nombre=<%=nombre_actual%>" title="<%=nombre_actual%>"><%=formato(nombre_actual,40)%></a>&nbsp;</td>
					<td class="celda_risctox" nowrap><% if instr(objRecordset("cas"),"-XX-")=0 then response.write formato(objRecordset("cas"),300)%>&nbsp;</td>
					<td class="celda_risctox" nowrap>
					<% 
					   if trim(ucase(objRecordset("fuente")))="NS" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=24','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>NS</a>"
					   end if
					   if trim(ucase(objRecordset("fuente")))="UE1" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=25','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>UE1</a>"
					   end if
					   if trim(ucase(objRecordset("fuente")))="UE2" then 
						response.write "<a onclick=window.open('ver_definicion.asp?id=26','def','width=300,height=200,scrollbars=yes,resizable=yes') style='cursor:hand'>UE2</a>"
					   end if %>&nbsp;</td>
				</tr>
<%				objRecordset.movenext
				i = i+1
				loop
				end if %>							
				<tr><td class=celda_risctox colspan=4>&nbsp;</td></tr>
<% 				if objRecordset.Pagecount>1 then%>
				<tr><td class=texto colspan=4 align="center">Hay <%=registros%> sustancias que cumplen los criterios de búsqueda. Se muestran sólo <%=registrosporpagina%> por página.</td></tr>
				<tr><td class=texto colspan=4 align="center">
<%				if Clng(Session("pagina")) > 1 then %>
				<a href="risctox_dis.asp?ordenacion=<%=ordenacion%>&buscar=<%=buscar%>&pag=<%=Clng(Session("pagina"))-1%>">anterior&nbsp;&lt;&lt;</a>&nbsp;&nbsp;&nbsp;
<%				end if %>
				Página&nbsp;<%=Session("pagina")%>&nbsp;de&nbsp;<%=objRecordset.Pagecount%>
<%				if Clng(Session("pagina")) < objRecordset.Pagecount then %>
				&nbsp;&nbsp;&nbsp;<a href="risctox_dis.asp?ordenacion=<%=ordenacion%>&buscar=<%=buscar%>&pag=<%=Clng(Session("pagina"))+1%>">&gt;&gt;&nbsp;siguiente</a>
<%				end if %>
				<form name="form_pag">Cambiar a la página:&nbsp;<input type="text" size="2" maxlenght="2" class="campo" value="<%=Session("pagina")%>" name="pag_ir">&nbsp;<input type="button" class="boton" value="ir" onclick="cambia_pag()"></form>
<%				end if %>				
				</td></tr>				
				</table>

				<% if registros>0 then %>
				<form name="imprimir" action="listado_imprimir.asp" method="POST" target="_blank">
				<input type="hidden" name="campo1" value="nombre">
				<input type="hidden" name="campo2" value="cas">
				<input type="hidden" name="campo3" value="fuente">
				<input type="hidden" name="nombre1" value="nombre/sustancia">
				<input type="hidden" name="nombre2" value="cas">
				<input type="hidden" name="nombre3" value="fuente">
				<input type="hidden" name="sql" value="<%=sql%>">
				<input type="hidden" name="buscar" value="<%=buscar%>">
				<input type="hidden" name="titulo" value="Listado de disruptores endocrinos">
				<p align="center"><input type="submit" class="boton" value="imprimir listado completo"></p>
				</form>
				<% end if %>

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
<script>
function cambia_pag()
{	if (form_pag.pag_ir.value><%=objRecordset.Pagecount%>)
	{ alert('Sólo se puede ir hasta la página <%=objRecordset.Pagecount%>'); }
	else
	{ location.href='risctox_dis.asp?ordenacion=<%=ordenacion%>&buscar=<%=buscar%>&pag='+form_pag.pag_ir.value; }
}
</script>
