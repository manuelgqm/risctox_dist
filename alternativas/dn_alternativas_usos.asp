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
<table width="100%" border="0">
<tr>
<td></td>
<td align='right'><input type="button" name="volver" class="boton" value="Volver a la portada de Alternativas" onClick="window.location='./index.asp';"></td>
</tr>
</table>
<p class=titulo3>Usos / Productos</p>

<%
if request("letra")="" then
	letra="A"
else
	letra=ucase(EliminaInyeccionSQL(request("letra")))
end if

sqll="SELECT DISTINCT LEFT(nombre, 1) AS letter FROM dn_risc_usos order by letter"
Set rstGetString=objConnection2.Execute(sqll)
if not rstGetString.eof then
	lista = rstGetString.GetString
	lista=ucase(lista)
end if
rstGetString.Close
Set rstGetString = Nothing

response.write "<p class='titulo3' align='center'>"
for i=65 to 90
	if i=79 then response.write hayresultados("Ñ")
	response.write hayresultados(chr(i))
next
response.write "</p>"
%>
<h2 class=titulo3><%=letra%></h2>
<%
	sqlmiletra="select id, nombre from dn_risc_usos where nombre like '" &letra& "%' order by nombre"
	set rstl=objConnection2.Execute(sqlmiletra)
	if rstl.eof then
		response.write "<p align='center'><strong>No hay resultados que comiencen con esta letra (" &letra& ")</strong></p>"
	else
		arrayDatos=rstl.getrows
		for contadorFilas=0 to ubound(arrayDatos,2)
			tablares=tablares& "<tr><td class='celda_risctox'><a href='dn_alternativas_ficha_uso.asp?id=" &arrayDatos(0,contadorFilas)& "'>" &arrayDatos(1,contadorFilas)& "</a></td><tr>"
		next

		'iniciotabla="<table class='tabla3' width='90%' align='center' border='0' cellpadding='4' cellspacing='0'> <tr><td class='subtitulo3'>	<table width='100%' align='center'><tr><td>"
		'iniciotabla=iniciotabla& "Alternativas &nbsp;</td><td align=right><!-- <img src='imagenes/ico_alt_procesos.gif'> -->"
		'iniciotabla=iniciotabla& "</td></tr></table></td> </tr>"
		tablares="<table class='tabla3' width='90%' align='center' border='0' cellpadding='4' cellspacing='0'>" &tablares& "</table>"
	end if
	rstl.close
	set rstl=nothing
	
	response.write tablares & "<br clear='all' />"
%>


<%
function hayresultados(letra)

	if instr(lista,letra) then letra = "<a href='dn_alternativas_usos.asp?letra="&letra&"'>" &letra& "</a>"
		
	hayresultados=letra & " &nbsp;"

end function
%>
		



<!-- ############ FIN DE CONTENIDO ################## -->
<!--#include file="spl_pie.inc.asp"-->

<%
cerrarconexion
%>
