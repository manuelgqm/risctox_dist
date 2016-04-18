<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->

<!--#include file="dn_restringida.asp"-->

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

<td><p class=campo>Est&aacute;s en: <a href=index.asp?idpagina=550>Plataforma prevención de riesgo químico</a>&nbsp;&gt;&nbsp;<a href="dn_alternativas_portada.asp">BBDD Alternativas</a> &gt; Sectores</p></td>

<td><input type="button" name="volver" class="boton" value="Volver a la portada de Alternativas" onclick="window.location='dn_alternativas_portada.asp';"></td>

</tr>

</table>




<p class=titulo3>Sectores</p>

<%
letra = request("letra")
letra = EliminaInyeccionSQL(letra)
if letra="" then
	letra="A"
else
	letra=ucase(letra)
end if

sqll="SELECT DISTINCT LEFT(nombre, 1) AS letter FROM dn_alter_sectores order by letter"
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
	sqlmiletra="select id, nombre from dn_alter_sectores where nombre like '" &letra& "%' order by nombre"
	set rstl=objConnection2.Execute(sqlmiletra)
	if rstl.eof then
		response.write "<p align='center'><strong>No hay resultados que comiencen con esta letra (" &letra& ")</strong></p>"
	else
		arrayDatos=rstl.getrows
		for contadorFilas=0 to ubound(arrayDatos,2)
			tablares=tablares& "<tr><td class='celda_risctox'><a href='dn_alternativas_ficha_sector.asp?id=" &arrayDatos(0,contadorFilas)& "'>" &arrayDatos(1,contadorFilas)& "</a></td><tr>"
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

	if instr(lista,letra) then letra = "<a href='dn_alternativas_sectores.asp?letra="&letra&"'>" &letra& "</a>"
		
	hayresultados=letra & " &nbsp;"

end function
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
cerrarconexion
%>


