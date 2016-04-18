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
            			<tr class="textmenusup"><td class=textmenusup>Formulario</td>
          		</table>
			</div>
			
			<div class="textsubmenu" id="submenusup">
			<table width="100%"  border="0" cellspacing="4" cellpadding="0">
            			<tr>
              				<td width="100%" valign="top">Est&aacute;s en: buscador</td>
            			</tr>
          		</table>
			</div>
			
			<div id="cuerpo">
					<font class="texto"><br>&nbsp;</font>
					<% if request("buscar")<>"" then
			   		   sql = "SELECT titulo,idpagina,numeracion FROM WEBISTAS_PAGINAS WHERE numeracion LIKE 'AI%' AND (titulo LIKE '%"&EliminaInyeccionSQL(request("buscar"))&"%' OR pagina LIKE '%"&EliminaInyeccionSQL(request("buscar"))&"%') ORDER BY numeracion"
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
						<a href="index.asp?idpagina=<%=objRecordset("idpagina")%>"><%=objRecordset("titulo")%></a></p>
					<%   objRecordset.movenext
			   		     loop
			   		   else %>
			   		   <p class="texto">No hay nigún resultado para esta búsqueda</p>
  					   <p>&nbsp;</p>
					   <p>&nbsp;</p>
					   <p>&nbsp;</p>
					   <p>&nbsp;</p>

			   		<% end if 
			   		   else %>
					  <p>&nbsp;</p>
					  <p>&nbsp;</p>
					  <p>&nbsp;</p>
					  <p>&nbsp;</p>
					  <p>&nbsp;</p>
			   		<% end if %>
					<p>&nbsp;</p>

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
