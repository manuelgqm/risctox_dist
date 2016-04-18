<!--#include file="EliminaInyeccionSQL.asp"-->
<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ECOinformas: buscador</title>
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
<body>
<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
			<div id="encabezado_nuevo1">
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
			<div id="menusup1">
			<table border="0" cellspacing="5" cellpadding="0">
            			<tr class="textmenusup"><td class=textmenusup>Mapa de la web y buscador</td>
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
            			<tr><td align="right">Usuari<%=usuario_sexo%> identificad<%=usuario_sexo%>:&nbsp;<%=usuario%></td></tr>
          		</table>
			</div>
       			<% end if %>
			
			<div id="cuerpo">
					&nbsp;
					<form name="form_busqueda" action="busqueda.asp">
					<p class="texto">Escribe el término a buscar:&nbsp;<input type="text" class="campo" value="<%=EliminaInyeccionSQL(request("buscar"))%>" name="buscar">&nbsp;
					<input type="submit" value="BUSCAR" class="boton"></p>
					</form>
				<% if EliminaInyeccionSQL(request("buscar"))<>"" then
			   		   sql = "SELECT titulo,idpagina,numeracion,tipo FROM WEBISTAS_PAGINAS WHERE numeracion LIKE 'AI%' AND (tipo=2 or tipo=7) AND visible=1"
			   		   if EliminaInyeccionSQL(request("buscar"))<>"" then sql = sql & " AND (titulo LIKE '%"&EliminaInyeccionSQL(request("buscar"))&"%' OR pagina LIKE '%"&EliminaInyeccionSQL(request("buscar"))&"%') "
			   		   sql = sql & " ORDER BY numeracion"
					   set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   		   set objRecordset = OBJConnection.Execute(sql)
		   	   		   if not objRecordset.eof then
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
						<% if objRecordset("tipo")=7 then response.write "&nbsp;<img src='imagenes\candado.gif' border=0 alt='Página restringida'>" %></p>
<%   					     objRecordset.movenext
			   		     loop
			   		   else %>
			   		   <p class="texto">No hay ningún resultado para esta búsqueda</p>
  					   <p>&nbsp;</p>
					   <p>&nbsp;</p>
					   <p>&nbsp;</p>
					   <p>&nbsp;</p>

			   		<% end if 
			   	else %>
					<table cellspacing=0 cellpadding=5 width="90%">
<%			   		   sql = "SELECT titulo,idpagina,numeracion,tipo FROM WEBISTAS_PAGINAS WHERE numeracion LIKE 'AI%' AND len(numeracion)>2 AND (tipo=2 or tipo=7) AND visible=1"
			   		   sql = sql & " ORDER BY numeracion"
					   set objRecordset = Server.CreateObject ("ADODB.Recordset")
		   	   		   set objRecordset = OBJConnection.Execute(sql)
		   	   		   if not objRecordset.eof then
		   	   		     r = 0
		   	   		     do while not objRecordset.eof
		   	   		     numeracion = objRecordset("numeracion")
		   	   		     if len(numeracion)>3 then
		   	   		     	r = r+1 %>
						<tr><td class="texto">
						<% 
              				     		for i=2 to len(numeracion)-3
              				   			'sql2 = "SELECT titulo,idpagina FROM WEBISTAS_PAGINAS WHERE numeracion='"&mid(numeracion,1,2+i)&"'" 
			   	   	   			'set objRecordset2 = Server.CreateObject ("ADODB.Recordset")
		   	   	           			'set objRecordset2 = OBJConnection.Execute(sql2)
		   	   	           			response.write "&nbsp;&gt;&nbsp;&nbsp;&nbsp;"
              				     		next %>
							<a href="index.asp?idpagina=<%=objRecordset("idpagina")%>"><%=objRecordset("titulo")%></a>
							<% if objRecordset("tipo")=7 then response.write "&nbsp;<img src='imagenes\candado.gif' border=0 alt='Página restringida'>" %>
							</td></tr>
              				  <%  else %>
              				     	</table>
              				     	<br>&nbsp;<table class="tabla<%=asc(mid(numeracion,3,1))-64%>" width="90%" cellpadding=5 align="center">
              				     	<tr><td><a href="index.asp?idpagina=<%=objRecordset("idpagina")%>"><img src="imagenes/icono<%=asc(mid(numeracion,3,1))-64%>.gif" border=0></a></td></tr>
					  <%  end if %>
						
					<%   objRecordset.movenext
			   		     loop
			   		   end if %>
			   		   </table>
			   	<% end if %>
					<p align=center class="texto"><img src='imagenes/candado.gif' border=0 alt='Página restringida'>&nbsp;Páginas restringidas. Se requiere clave para acceder que se puede solicitar pulsando <a href="formulario2.asp">aquí</a></p>

			<map name="Map1" id="Map1">
            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundación Biodiversidad" />
            		<area shape="rect" coords="400,38,546,104" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="547,38,701,104" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
			<img src="imagenes/pie1.jpg" width="708" border="0" usemap="#Map1">

    			</div>
    		</div>
		<div id="sombra_abajo"></div>
	</div>
</body>
</html>
