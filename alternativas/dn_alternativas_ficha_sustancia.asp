<!--#include file="../dn_restringida.asp"-->
<!--#include file="../dn_conexion.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->
<!--#include file="../dn_restringida.asp"-->

<%
' Inicialmente no hay errores..
errores = ""

' Cogemos el id de la sustancia elegida y traemos sus datos
id_sustancia = EliminaInyeccionSQL(request("id_sustancia"))

' Nombre
sql="SELECT nombre FROM dn_risc_sustancias WHERE id="&id_sustancia
set objRst=objConnection2.execute(sql)
if(objRst.eof) then
	errores="No se ha encontrado la sustancia indicada"
else
	nombre = objRst("nombre")
end if
objRst.close()
set objRst=nothing

' Sinonimos
sinonimos = dameSinonimos(id_sustancia)
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: BBDD Alternativas</title>
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
<link rel="stylesheet" type="text/css" href="dn_estilos.css">

<script type="text/javascript" src="scripts/moo/prototype.lite.js"></script>
<script type="text/javascript" src="scripts/moo/moo.fx.js"></script>
<script type="text/javascript" src="scripts/moo/moo.fx.pack.js"></script>
<script type="text/javascript" src="scripts/moo/moo.ajax.js"></script>

<script language="JavaScript">
function uso_change()
{
	// Cogemos los campos para la búsqueda
	id_uso = document.getElementById("uso").options[document.getElementById("uso").selectedIndex].value;
	id_sustancia=<%=id_sustancia%>;

	// Mostramos mensaje "Buscando..."
	document.getElementById("sustancias_por_uso").innerHTML="<div class='mensaje_ajax'><img src='imagenes/progress.gif' hspace='5' align='absmiddle'/><strong>Buscando sustancias alternativas...</strong> Por favor, espere.</div>";

	// Enviamos consulta AJAX
	new ajax('dn_alternativas_ficha_ajax.asp', {postBody: 'id_uso='+id_uso+"&id_sustancia="+id_sustancia, update: $('sustancias_por_uso'), onComplete: uso_change_completed});
}

// #####################################################################

function uso_change_completed()
{

}
</script>

</head>
<body>
<div id="contenedor">
	<div id="sombra_arriba"></div>
  	<div id="sombra_lateral">
		<div id="caja">
		<!--#include file="../dn_cabecera.asp"-->
		<div id="texto">
			
<div class="texto">
<!-- ################ CONTENIDO ###################### -->


<%
if (errores<>"") then
	response.write errores
else
%>
	<table width="100%" border="0">
<tr>
<td></td>
<td align='right'><input type="button" name="volver" class="boton" value="Volver a la portada de Alternativas" onClick="window.location='./index.asp';"></td>
</tr>
</table>
	<!-- Datos de sustancia -->
	<p class=titulo3>Datos de la sustancia</p>
	<table width="100%" class="tabla3">
		<tr>
			<td class="titulo3" align="right">
				Nombre:
			</td>
			<td class="texto">
				<b><a href="http://www.istas.net/risctox/dn_risctox_ficha_sustancia.asp?id_sustancia=<%=id_sustancia%>" title="Ver ficha en Risctox" /><%=espaciar(nombre)%></a></b>
			</td>
		</tr>
		<%
		if (sinonimos<>"") then
		%>
			<tr>
				<td class="titulo3" align="right">
					Sinónimos:
				</td>
				<td class="texto">
					<%=sinonimos%>
				</td>
			</tr>
		<%
		end if ' hay sinonimos?
		%>    
	</table>

<!-- Sustancias alternativas asociadas a través del uso -->
<% muestraUsos id_sustancia %>

<!-- Ficheros de alternativas asociados -->
<% muestraFicheros id_sustancia, "Documentos de alternativas" %>

<!-- Casos prácticos asociados -->
<% muestraFicheros id_sustancia, "Casos prácticos" %>

<%
end if ' Comprobación de errores
%>





<!-- ############ FIN DE CONTENIDO ################## -->

<br>
Esta página ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundación de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a><br>

				
				</div>
				<p>&nbsp;</p>
			</div>
			
			
			<img src="../imagenes/pie_risctox.gif" width="708" border="0">
      			
    			
    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>
</body>
</html>

<%
cerrarconexion
%>


<%
' Funciones auxiliares

' ##########################################################################
sub muestraUsos(byval id_sustancia)
	' Muestra un desplegable con los usos en los que la sustancia indicada aparece asociada como tóxica
	sql = "SELECT dn_risc_sustancias_por_usos.id_uso AS id_uso, dn_risc_usos.nombre AS nombre FROM dn_risc_sustancias_por_usos INNER JOIN dn_risc_usos ON dn_risc_sustancias_por_usos.id_uso = dn_risc_usos.id WHERE (dn_risc_sustancias_por_usos.toxico = 1) AND (dn_risc_sustancias_por_usos.id_sustancia="&id_sustancia&") ORDER BY nombre"
	'response.write "<br/>"&sql
	set objRst = objConnection2.execute(sql)

	' Nos quedamos sólo con aquellos usos que también tienen asociada alguna sustancia como no tóxica
	' Construimos una cadena de id_uso separada por comas
	ids_usos = ""
	do while (not objRst.eof)
		' Para cada uso, contamos el numero de alternativas menos toxicas que tiene
		sql2 = "SELECT count(*) AS num FROM dn_risc_sustancias_por_usos WHERE id_uso = "&objRst("id_uso")&" AND toxico = 0"
		'response.write "<br/>"&sql2
		set objRst2=objConnection2.execute(sql2)
		do while (not objRst2.eof)
			if (objRst2("num") > 0) then
				if (ids_usos = "") then
					ids_usos = objRst("id_uso")
				else
					ids_usos = ids_usos & ", " & objRst("id_uso")
				end if	
				'response.write "<br/>"&ids_usos
			end if
			objRst2.movenext
		loop
		objRst2.close()
		set objRst2=nothing
		objRst.movenext
	loop
	objRst.close()
	set objRst=nothing

	' Si ids_usos tiene algo, es que hay usos con alternativas.

	if (ids_usos <> "") then
		sql3="SELECT DISTINCT dn_risc_sustancias_por_usos.id_uso AS id_uso, dn_risc_usos.nombre AS nombre FROM dn_risc_sustancias_por_usos INNER JOIN dn_risc_usos ON dn_risc_sustancias_por_usos.id_uso = dn_risc_usos.id WHERE (dn_risc_sustancias_por_usos.id_uso IN ("&ids_usos&")) ORDER BY nombre"
		'response.write "<br/>"&sql3
		set objRst3=objConnection2.execute(sql3)
%>

<p class=titulo3>Sustancias alternativas</p>
<p>Selecciona el uso que tiene la sustancia que quieres sustituir, para obtener una relación de sustancias alternativas.</p>
<form name="form_alternativas">
<table width="100%" class="tabla3">
	<tr>
		<td class="titulo3"align="right">
			Uso:
		</td>
		<td class="texto">
			<select name="uso" id="uso" onChange="uso_change();" />
				<option value="0">Seleccione el uso</option>
<%
			do while (not objRst3.eof)
%>
				<option value="<%=objRst3("id_uso")%>"><%=objRst3("nombre")%></option>
<%
				objRst3.movenext
			loop
			objRst3.close()
			set objRst3=nothing
%>	
			</select>
      &nbsp;alternativo al <%= nombre %>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><div id="sustancias_por_uso"></div></td>
	</tr>
</table>
</form>
<%
	else
		'response.write "<p>No se han encontrado usos para esta sustancia</p>"
	end if
end sub

' ##########################################################################

sub muestraFicheros(byval id_sustancia, byval tipo)
	' Muestra los ficheros de alternativas asociados a la sustancia
	' para el tipo indicado ("alternativas" o "casos")

	if (tipo = "Casos prácticos") then
		filtro = " (dn_alter_ficheros.tema = 'Caso práctico' OR dn_alter_ficheros.tema = 'Casos prácticos') AND "
  else
		filtro = " (dn_alter_ficheros.tema <> 'Caso práctico' AND dn_alter_ficheros.tema <> 'Casos prácticos') AND "
	end if

	sql = "SELECT id_fichero, titulo, tema FROM dn_alter_ficheros_por_sustancias INNER JOIN dn_alter_ficheros ON dn_alter_ficheros_por_sustancias.id_fichero = dn_alter_ficheros.id WHERE "& filtro &" dn_alter_ficheros_por_sustancias.id_sustancia="&id_sustancia
	set objRst = objConnection2.execute(sql)
	if (not objRst.eof)then
%>
	<p class=titulo3><%=tipo%></p>
  <p><%=tipo%> relacionados con el <%= nombre %>:</p>
	<table width="100%" class="tabla3">
		<tr>
			<td>&nbsp;</td>
			<td class="texto">
				<ul>
	<%
				' Evitamos mostrar titulos repetidos
				titulo_antiguo = ""
				do while (not objRst.eof)
					if (objRst("titulo") <> titulo_antiguo) then
	%>
						<li><a href="dn_alternativas_ficha_fichero.asp?id_fichero=<%=objRst("id_fichero")%>"><%=objRst("titulo")%></a></li>
	<%
					end if
					titulo_antiguo = objRst("titulo")
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
