<!--#include file="dn_restringida.asp"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->
<!--#include file="dn_restringida.asp"-->
<!--#include file="../EliminaInyeccionSQL.asp"-->

<%
' Inicialmente no hay errores..
errores = ""

' Cogemos el id de la sustancia elegida y traemos sus datos
id_fichero = request("id_fichero")
id_fichero = EliminaInyeccionSQL(id_fichero)

sql="SELECT * FROM dn_alter_ficheros WHERE id="&id_fichero
set objRst=objConnection2.execute(sql)
if(objRst.eof) then
	errores="No se ha encontrado el fichero indicado"
end if
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: risctox</title>
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

<link rel="stylesheet" type="text/css" href="estructura.css">
<link rel="stylesheet" type="text/css" href="dn_estilos.css">

</head>
<body>
<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
		<!--#include file="dn_cabecera.asp"-->
		<div id="texto">
			
<div class="texto">
<!-- ################ CONTENIDO ###################### -->

<table width="100%" border="0">
<tr>
<td><p class=campo>Est&aacute;s en: <a href=index.asp?idpagina=550>Plataforma prevención de riesgo químico</a>&nbsp;&gt;&nbsp;<a href="dn_alternativas_portada.asp">BBDD Alternativas</a>&nbsp;&gt;&nbsp;Fichero de alternativas</p></td>
<td><input type="button" name="volver" class="boton2" value="Volver a la portada de Alternativas" onclick="window.location='dn_alternativas_portada.asp';"></td>
</tr>
</table>

<%
if (errores<>"") then
	response.write errores
else
%>

	<!-- Datos del fichero -->
	<p class=titulo3>Base de datos de alternativas de sustitución de productos con riesgo tóxico</p>
	<table width="100%" class="tabla3">
		<% if (objRst("titulo") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Alternativa:</td>
			<td class="texto"><b><%=objRst("titulo")%></b></td>
		</tr>  
		<% end if %>

		<% if (objRst("resumen") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Resumen:</td>
			<td class="texto"><%=objRst("resumen")%></td>
		</tr>
		<% end if %>
		
        <% if (objRst("criterios_aceptacion") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Criterios de aceptacion:</td>
			<td class="texto"><%=objRst("criterios_aceptacion")%></td>
		</tr>
		<% end if %>
        
        
		<% if (objRst("archivo") <> "") then %>
    <%
    if (session("modo") = "pruebas") then
      rutaficheros="web/pruebas/gestion/estructuras"
    elseif (session("modo") = "produccion") then
      rutaficheros="gestion/estructuras"
    end if
    %>
		<tr>
			<td class="titulo3" align="right">Archivo:</td>
			<td class="texto"><a href="http://www.istas.net/risctox/<%= rutaficheros %>/<%=objRst("archivo")%>">Descargar archivo</a> (500 Kb)</td>
		</tr>
		<% end if %>

		<% if (objRst("direccion_internet") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Enlace:</td>
			<td class="texto"><a href="<%=objRst("direccion_internet")%>"><%= left(objRst("direccion_internet"), 80) %></a></td>
		</tr> 
		<% end if %>

		<% if (objRst("idioma") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Idioma:</td>
			<td class="texto"><%=objRst("idioma")%></td>
		</tr> 
		<% end if %>

		<% if (objRst("autor") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Autor:</td>
			<td class="texto"><%=objRst("autor")%></td>
		</tr> 
		<% end if %>

		<% if (objRst("lugar") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Lugar:</td>
			<td class="texto"><%=objRst("lugar")%></td>
		</tr> 
		<% end if %>

		<% if (objRst("publicacion") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Publicación:</td>
			<td class="texto"><%=objRst("publicacion")%></td>
		</tr> 
		<% end if %>

		<% if (objRst("coleccion") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Colección:</td>
			<td class="texto"><%=objRst("coleccion")%></td>
		</tr> 
		<% end if %>

		<% if (objRst("descripcion_fisica") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Descripción:</td>
			<td class="texto"><%=objRst("descripcion_fisica")%></td>
		</tr> 
		<% end if %>

		<% if (objRst("numero_normalizado") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Número normalizado:</td>
			<td class="texto"><%=objRst("numero_normalizado")%></td>
		</tr> 
		<% end if %>

		<% if (objRst("notas") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Notas:</td>
			<td class="texto"><%=objRst("notas")%></td>
		</tr> 
		<% end if %>

		<% if (objRst("soporte") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Soporte:</td>
			<td class="texto"><%=objRst("soporte")%></td>
		</tr> 
		<% end if %>

		<% if (objRst("fecha_actualizacion") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Fecha de actualización:</td>
			<td class="texto"><%=objRst("fecha_actualizacion")%></td>
		</tr>
		<% end if %>

		<% if (objRst("fecha_consulta") <> "") then %>
		<tr>
			<td class="titulo3" align="right">Fecha de consulta:</td>
			<td class="texto"><%=objRst("fecha_consulta")%></td>
		</tr>
		<% end if %>
	</table>

<%
end if ' Comprobación de errores
%>





<!-- ############ FIN DE CONTENIDO ################## -->



<br>
Esta página ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundación de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a><br>

				
				</div>
				<p>&nbsp;</p>
			</div>
			
			
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
			<img src="imagenes/pie3.jpg" width="708" border="0" usemap="#Map3">
      			
    			
    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>
</body>
</html>

<%
objRst.close()
set objRst=nothing

cerrarconexion
%>


<%
' Funciones auxiliares

' ##########################################################################
sub muestraSectores(byval id_fichero)
	' Muestra la lista de sectores relacionados con el fichero
	sql = "SELECT id_sector, nombre FROM dn_alter_ficheros_por_sectores INNER JOIN dn_alter_sectores ON dn_alter_ficheros_por_sectores.id_sector = dn_alter_sectores.id WHERE dn_alter_ficheros_por_sectores.id_fichero = "&id_fichero
	set objRst = objConnection2.execute(sql)
	if (not objRst.eof) then
%>

<p class=titulo3>Sectores relacionados</p>
<table width="100%" class="tabla3">
	<tr>
		<td class="titulo3" align="right">&nbsp;
			
		</td>
		<td class="texto">
			<ul>
<%
			do while (not objRst.eof)
%>
				<li><a href="dn_alternativas_ficha_sector.asp?id_sector=<%=objRst("id_sector")%>"><%=objRst("nombre")%></a></li>
<%
				objRst.movenext
			loop
%>	
			</ul>
		</td>
	</tr>
</table>
<%
	else
		'response.write "<p>No se han encontrado sectores para este fichero</p>"
	end if
end sub

' ##########################################################################
sub muestraProcesos(byval id_proceso)
	' Muestra la lista de procesos relacionados con el fichero
	sql = "SELECT id_proceso, nombre FROM dn_alter_ficheros_por_procesos INNER JOIN dn_alter_procesos ON dn_alter_ficheros_por_procesos.id_proceso = dn_alter_procesos.id WHERE dn_alter_ficheros_por_procesos.id_fichero = "&id_fichero
	set objRst = objConnection2.execute(sql)
	if (not objRst.eof) then
%>

<p class=titulo3>Procesos relacionados</p>
<table width="100%" class="tabla3">
	<tr>
		<td class="titulo3" align="right">&nbsp;
			
		</td>
		<td class="texto">
			<ul>
<%
			do while (not objRst.eof)
%>
				<li><a href="dn_alternativas_ficha_proceso.asp?id_proceso=<%=objRst("id_proceso")%>"><%=objRst("nombre")%></a></li>
<%
				objRst.movenext
			loop
%>	
			</ul>
		</td>
	</tr>
</table>
<%
	else
		'response.write "<p>No se han encontrado procesos para este fichero</p>"
	end if
end sub

' ##########################################################################
sub muestraUsos(byval id_fichero)
	' Muestra la lista de usos relacionados con el fichero
	sql = "SELECT id_uso, nombre FROM dn_alter_ficheros_por_usos INNER JOIN dn_risc_usos ON dn_alter_ficheros_por_usos.id_uso = dn_risc_usos.id WHERE dn_alter_ficheros_por_usos.id_fichero = "&id_fichero
	set objRst = objConnection2.execute(sql)
	if (not objRst.eof) then
%>

<p class=titulo3>Usos relacionados</p>
<table width="100%" class="tabla3">
	<tr>
		<td class="titulo3" align="right">&nbsp;
			
		</td>
		<td class="texto">
			<ul>
<%
			do while (not objRst.eof)
%>
				<li><a href="dn_alternativas_ficha_uso.asp?id_uso=<%=objRst("id_uso")%>"><%=objRst("nombre")%></a></li>
<%
				objRst.movenext
			loop
%>	
			</ul>
		</td>
	</tr>
</table>
<%
	else
		'response.write "<p>No se han encontrado usos para este fichero</p>"
	end if
end sub

' ##########################################################################

sub muestraFicheros(byval id_sustancia, byval tipo)
	' Muestra los ficheros de alternativas asociados a la sustancia
	' para el tipo indicado ("alternativas" o "casos")

	if (tipo = "Casos prácticos") then
		filtro = " dn_alter_ficheros.tema = 'Caso práctico' AND "
	end if

	sql = "SELECT id_fichero, titulo FROM dn_alter_ficheros_por_sustancias INNER JOIN dn_alter_ficheros ON dn_alter_ficheros_por_sustancias.id_fichero = dn_alter_ficheros.id WHERE "& filtro &" dn_alter_ficheros_por_sustancias.id_sustancia="&id_sustancia
	set objRst = objConnection2.execute(sql)
	if (not objRst.eof)then
%>
	<p class=titulo3><%=tipo%></p>
	<table width="100%" class="tabla3">
		<tr>
			<td>&nbsp;</td>
			<td class="texto">
				<ul>
	<%
				do while (not objRst.eof)
	%>
					<li><a href="dn_alternativas_ficha_fichero.asp?id_fichero=<%=objRst("id_fichero")%>"><%=objRst("titulo")%></a></li>
	<%
					objRst.movenext
				loop
	%>	
				</ul>
			</td>
		</tr>
	</table>
<%
	else
		'response.write "<p>No hay ficheros asociados</p>"
	end if
end sub

%>
