<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_conexion.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->
<!--#include file="../dn_restringida.asp"-->

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
	id=EliminaInyeccionSQL(request("id"))
	set rstl=objConnection2.Execute("select nombre from rq_residuos where id =" &id)
	nombre=rstl("nombre")
	rstl.close
	set rstl=nothing
%>

	<!-- Datos de sustancia -->
	<p class=titulo3><%=nombre%></p>

<!-- FICHEROS DE ALTERNATIVAS 
 ficheros asociados al uso.  -->

	
<%
	set rstl=objConnection2.Execute("select dn_alter_ficheros.id, dn_alter_ficheros.titulo, dn_alter_ficheros.num_alternativa from dn_alter_ficheros_por_residuos JOIN dn_alter_ficheros ON dn_alter_ficheros_por_residuos.id_fichero=dn_alter_ficheros.id where dn_alter_ficheros.tema<>'Casos prácticos' and dn_alter_ficheros_por_residuos.id_residuo =" &id & "ORDER BY dn_alter_ficheros.titulo")
	if rstl.eof then
		response.write "No se encontraron ficheros asociados a este residuo<br /><br /><br />"

	else
%>
<table width="100%" border="0">
<tr>
<td></td>
<td align='right'><input type="button" name="volver" class="boton" value="Volver a la portada de Alternativas" onClick="window.location='./index.asp';"></td>
</tr>
</table>
<p class=titulo3>Documentos de alternativas</p>
	<table width="100%" class="tabla3">
		<tr>
			<td class="texto">
<%
		response.write "<ul>"
		nums_alternativas = ""
		do while not rstl.eof
			' Mostramos el documento si su número de alternativa no había salido antes
			' Lo buscamos siempre con una coma detrás
			'response.write instr(nums_alternativas, rstl(2)&",") & "<br/>"
			if (instr(nums_alternativas, rstl(2)&",") = 0) then
				response.write "<li><a href='dn_alternativas_ficha_fichero.asp?id_fichero=" &rstl(0)& "'>" &rstl(1)& "</a></li>"
				' Apuntamos el num_alternativa en la lista para no repetir
				if (nums_alternativas <> "") then
					nums_alternativas = nums_alternativas & rstl(2) & ","
				else
					nums_alternativas = rstl(2)&","
				end if
			end if

		rstl.movenext
		loop
		response.write "</ul>"

		rstl.close
		set rstl=nothing
%>
		</td>
		</tr>
	</table>
<%
	end if
%>
			

<!-- CASOS PRACTICOS asociados -->

	
<%
	' Ignoramos los que tengan numero de alternativa mostrado en documentos

	set rstl=objConnection2.Execute("select dn_alter_ficheros.id, dn_alter_ficheros.titulo, dn_alter_ficheros.num_alternativa from dn_alter_ficheros_por_residuos JOIN dn_alter_ficheros ON dn_alter_ficheros_por_residuos.id_fichero=dn_alter_ficheros.id where dn_alter_ficheros.tema='Casos prácticos' and dn_alter_ficheros_por_residuos.id_residuo =" &id & " AND dn_alter_ficheros.num_alternativa NOT IN ("&nums_alternativas&"-1) ORDER BY dn_alter_ficheros.titulo")
	if rstl.eof then
		'response.write "No se encontraron Casos prácticos"

	else
%>
<p class=titulo3>Casos prácticos</p>
	<table width="100%" class="tabla3">
		<tr>
			<td class="texto">
<%
		response.write "<ul>"
		titulo_antiguo=""		
		do while not rstl.eof
			if (rstl(1) <> titulo_antiguo) then			
				response.write "<li><a href='dn_alternativas_ficha_fichero.asp?id_fichero=" &rstl(0)& "'>" &rstl(1)& "</a></li>"
				titulo_antiguo = rstl(1)
			end if
			rstl.movenext
		loop
		response.write "</ul>"
		rstl.close
		set rstl=nothing
%>
</td>
		</tr>
	</table>
<%
	end if
%>
			

<!-- ############ FIN DE CONTENIDO ################## -->
<!--#include file="spl_pie.inc.asp"-->

<%
cerrarconexion
%>
