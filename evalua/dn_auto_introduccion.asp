<!--#include file="dn_conexion.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: BBDD Eval�a lo que usas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multim�dia" />
<meta name="description" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Informaci�n, formaci�n y asesoramiento sobre medio ambiente para trabajadores de PYME" />
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
<img src=http://www.istas.net/recursos/IMG/ISTAS_01033.jpg align="right" hspace="20" vspace="10">

<p class=titulo3>Bienvenid@ a Eval�a y Compara lo que Usas</p>

<p>En esta secci�n te ofrecemos una herramienta automatizada, de f�cil utilizaci�n, que te permite:</p>

<p><ul><li><strong>Evaluar el riesgo que presentan los productos qu�micos que utilizas en tu empresa</strong>, a partir de la informaci�n existente en las Fichas de Datos de Seguridad (n�meros CAS y frases R).</li>
<li><strong>Comparar los riesgos que presentan diferentes productos</strong>, a fin de encontrar la alternativa m�s adecuada a las condiciones de uso que t� especifiques.</li></ul></p>

<p>En resumen, el objetivo es ayudarte en el proceso de sustituci�n de productos qu�micos peligrosos por aquellos que no sean tan nocivos para tu salud y para el medio ambiente.</p>

<p align=center>
<%
if session("id_ecogente")="" then
	session("redirigir_tras_autentificar")="dn_auto_introduccion.asp"
%>
<input type=button class=boton value="Identif�cate para poder usar la herramienta" onclick=location.href="acceso.asp">
<%
else
%>
<input type=button class=boton value="Ir a Herramienta de Autoevaluaci�n" onclick=location.href="dn_auto_portada.asp">
<%
end if
%>
</p>				
				</div>
				<p>&nbsp;</p>
			</div>
			
			
			<map name="Map1" id="Map1">
            		<area shape="rect" coords="307,38,399,104" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
            		<area shape="rect" coords="400,38,546,104" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="547,38,701,104" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>

			<map name="Map2" id="Map2">
            		<area shape="rect" coords="300,8,392,66" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
            		<area shape="rect" coords="393,8,539,66" href="http://www.istas.ccoo.es" target="_blank" alt="Instituto Sindical de Trabajo, Ambiente y Salud" />
      			<area shape="rect" coords="540,8,694,66" href="http://www.mtas.es/UAFSE/default.htm" target="_blank" alt="Fondo Social Europeo" />
      			</map>
      			<map name="Map3" id="Map3">
            		<area shape="rect" coords="300,18,392,80" href="http://www.fundacion-biodiversidad.es" target="_blank" alt="Fundaci�n Biodiversidad" />
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
