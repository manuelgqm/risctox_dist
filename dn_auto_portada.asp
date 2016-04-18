<!--#include file="dn_conexion.asp"-->

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
<img src=http://www.istas.net/recursos/IMG/ISTAS_01033.jpg align="right" hspace="20" vspace="10">

<p class=titulo3>¿Qué es Evalúa y Compara lo que Usas?</p>

<p>"Evalúa y Compara lo que Usas" es una herramienta que ayuda a valorar los riesgos para la salud y el medio ambiente de los productos químicos utilizados en los lugares de trabajo. Al permitir comparar los riesgos de varios productos, puede servir de ayuda en la búsqueda de alternativas que eviten o reduzcan el riesgo químico en las empresas.

<p>Es un recurso basado en el <a href="pdf/NTP_712.doc">Modelo de Columnas</a> desarrollado por el BIA (Instituto de Mutuas Profesionales Alemanas) y adaptado por el ISTAS.</p>

<p>NO ES UNA EVALUACIÓN EN PROFUNDIDAD, sino tan sólo una propuesta automatizada que realiza una valoración preliminar del riesgo de un producto y sus componentes, y que depende del nivel de fiabilidad de la información que se facilite.</p>

<p>No se pretende con ella un dictamen, tan sólo una orientación, un primer paso, sencillo y práctico, en el camino de la sustitución de sustancias peligrosas.</p>

<p align=center>
<%
if session("id_ecogente")="" then
	session("redirigir_tras_autentificar")="dn_auto_portada.asp"
%>
<input type=button class=boton value="Identifíquese para poder usar la herramienta" onclick=location.href="acceso.asp">
<%
else
%>
<input type=button class=boton value="Ir a Herramienta de Autoevaluación" onclick=location.href="dn_auto_herramienta.asp">
<%
end if
%>
</p><br>

<p class="titulo3">¿Cómo utilizarla?</p>
<p><strong>* Ficha cuestionario</strong>. Se rellena una ficha por cada producto que se quiera evaluar. La información necesaria se encuentra en las Fichas de Datos de Seguridad de los Productos (<a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=529">FDS</a>). Los iconos <img src="imagenes/fam_istas/help.gif" align="absmiddle"> proporcionan información sobre cómo completar cada apartado.<br/><br/>
Parte de la información sobre los componentes: nombre, números de identificación, Frases R se completa automáticamente si la sustancia se encuentra en la base de datos <a href="http://www.istas.net/ecoinformas/web/index.asp?idpagina=575">RISCTOX</a>.<br/><br/>
En el supuesto de que no se rellenaran todos los campos o consideras que la información no coincide con la que tu dispones, puedes completarla y modificarla. Una vez que hayas completado la ficha pulsa sobre el botón GUARDAR PRODUCTO.</p>

<p><strong>* Resultados</strong>. Para hacer la evaluación preliminar del producto, pulsa sobre EVALUAR/COMPARAR. La información que verás es una evaluación básica sobre los niveles de riesgo de cada una de las sustancias que componen el producto sobre el que realizas la consulta.</p>

<p><strong>* Comparación</strong>. Para comparar varios productos elige de tu LISTA DE PRODUCTOS los que quieres comparar y pulsa el botón EVALUAR/COMPARAR. Los productos y sustancias se comparan por columnas, esto es  por tipos de riesgo (toxicidad aguda; ecotoxicidad; etc.). Además, se deben tener en cuenta las condiciones de uso del producto. A la vista de los niveles de riesgo identificados por la herramienta deberás de optar por el producto o sustancia que presente los niveles más bajos.</p>

<p align="center">

<img src="imagenes/dn_ejemplo_evalua.jpg" alt="Ejemplo de tabla de evaluación" title="Ejemplo de tabla de evaluación"><br /><br />

<strong>SIEMPRE SE DEBE TOMAR EN CONSIDERACIÓN LA SITUACIÓN DE LA EMPRESA RESPECTO DE LAS CONDICIONES DE PREVENCIÓN Y GESTIÓN DEL RIESGO IMPLANTADAS.</strong></p>

<p class="titulo3">Por qué impulsar la sustitución</p>
<p>La <a href="pdf/EcoSustanciasDefinitivaLEX.pdf">normativa básica</a> de referencia en riesgo químico establece como prioridad la eliminación del riesgo, por lo que la sustitución, en tanto que técnica preventiva, resulta una prioridad cuando no una obligación (cancerígenos, mutágenos y algunos tóxicos para la reproducción). Además, es prioritario eliminar o sustituir todas las sustancias que debido a su peligrosidad intrínseca presentan un nivel de riesgo inaceptable, incluidas en la <a href="pdf/lista_negra_istas.pdf" target="_blank">lista negra de ISTAS</a>.</p>
<p>Al proponer a la empresa cualquier iniciativa de sustitución debemos concretar en primer lugar unos criterios para la búsqueda de alternativas. Después estableceremos unas etapas por las que avanzar en la materialización de la iniciativa. Puedes ampliar información en la <a href="http://www.istas.net/ecoinformas/web/abreenlace.asp?idenlace=2428">Guía para la sustitución de sustancias peligrosas</a>.</p>
<p align="center">
<%
if session("id_ecogente")="" then
	session("redirigir_tras_autentificar")="dn_auto_portada.asp"
%>
<input type=button class=boton value="Identifíquese para poder usar la herramienta" onclick=location.href="acceso.asp">
<%
else
%>
<input type=button class=boton value="Ir a Herramienta de Autoevaluación" onclick=location.href="dn_auto_herramienta.asp">
<%
end if
%>
</p>

<a id="adaptado_istas"></a>
<p class="titulo3">Notas sobre la adaptación de ISTAS</p>
<p>A diferencia del modelo de columnas, esta herramienta considera sustancias de muy alto riesgo de toxicidad crónica:</p>

<p>
	<ul>
		<li>Las sustancias cancerígenas C3 (R40) y mutágenas M3 (R68) según el RD 363/1995 y las sustancias cancerígenas 1, 2A y 2B según IARC.</li>
		<li>Las sustancias tóxicas para la reproducción: R60, R61, R62 y R63.</li>
		<li>Las sustancias bioacumulables (R33) y que se acumulan en la leche materna (R64).</li>
		<li>Las sustancias sensibilizantes, neurotóxicas y disruptores endocrinos.</li>
	</ul>
</p>

<p>Además, considera de muy alto riesgo para el medio ambiente las sustancias tóxicas, persistentes y bioacumulativas y los disruptores endocrinos.</p>

<p>La herramienta utiliza los listados de sustancias peligrosas de la base de datos RISCTOX elaborada por ISTAS.
<br/><br/></p>

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
