<!--#include file="../dn_conexion.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS Evalúa lo que usas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="Evalúa lo que usas" />
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
		<!--#include file="../dn_cabecera.asp"-->
		<div id="texto">
			
<div class="texto">
<img src=http://www.istas.net/recursos/IMG/ISTAS_01033.jpg align="right" hspace="20" vspace="10">

<p class=titulo3>¿Qué es Evalúa y Compara lo que Usas?</p>

<p>"Evalúa y Compara lo que Usas" es una herramienta para ayudar a valorar los riesgos para la salud y el medio ambiente de las sustancias y los productos químicos utilizados en los lugares de trabajo. Al permitir comparar los riesgos de varios productos, puede servir de ayuda en la búsqueda de alternativas que eviten o reduzcan el riesgo químico en las empresas. 
</p>
<p>
Es un recurso basado en el <a href='http://www.hvbg.de/e/bia/pra/spalte/index.html' target='_blank'>Modelo de Columnas</a> desarrollado por el BIA (Instituto de Mutuas Profesionales Alemanas) y adaptado por el ISTAS.
</p>
<p>
NO ES UNA EVALUACIÓN EN PROFUNDIDAD, sino tan sólo una propuesta automatizada que realiza una valoración preliminar del riesgo de un producto y sus componentes, y que depende del nivel de fiabilidad de la información que se facilite.
</p>
<p>
No se pretende con ella un dictamen, tan sólo una orientación, un primer paso, sencillo y práctico, en el camino de la sustitución de sustancias peligrosas.
</p>
<p>
La matriz permite:
<ul>
	<li>Evaluar el riesgo que presentan los productos químicos que utilizas en tu empresa, a partir de la información existente en las Fichas de Datos de Seguridad (números CAS y frases R).</li>
    <li>Comparar los riesgos que presentan diferentes productos, a fin de encontrar la alternativa más adecuada a las condiciones de uso que tú especifiques.</li>
</ul>
</p>
<p>
Incluye las siguientes variables: efectos agudos, efectos crónicos, ecotoxicidad, inflamabilidad y explosión, tipo de exposición y tipo de proceso de trabajo.
</p>
<p>
Permite clasificar cada una de las sustancias según los siguientes niveles de riesgo: muy alto, alto, medio, bajo y muy bajo. 
</p>
<p>
En aquellas situaciones de riesgo en las que coincida más de una sustancia química, tendremos que realizar una evaluación de la situación de riesgo que resuma los resultados parciales de la evaluación de cada una de las sustancias implicadas. 
</p>
<p>
Dado que en condiciones de multiexposición es probable que los efectos de cada una de las sustancias, se sumen o incluso multipliquen el resultado final, calificaremos le riesgo resultante de un producto, como mínimo, igual al nivel de riesgo encontrado para alguna de las sustancias que lo componen.  
</p>
<p>
Esta matriz no considera las concentraciones en las que se encuentran las sustancias en los preparados, lo que puede dar lugar a una calificación del riesgo de algunos productos superior a la establecida en las frases R que aparecen en su etiqueta o FDS.
</p>


<p align=center>
<%
if session("id_ecogente2")="" then
	'session("redirigir_tras_autentificar")="dn_auto_portada.asp"
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
<p><strong>* Ficha cuestionario</strong>. Se rellena una ficha por cada producto que se quiera evaluar. La información necesaria se encuentra en las Fichas de Datos de Seguridad de los Productos (<a href="http://www.istas.net/risctox/index.asp?idpagina=529">FDS</a>). Los iconos <img src="imagenes/fam_istas/help.gif" align="absmiddle"> proporcionan información sobre cómo completar cada apartado.<br/><br/>
Parte de la información sobre los componentes: nombre, números de identificación, Frases R se completa automáticamente si la sustancia se encuentra en la base de datos <a href="http://www.istas.net/risctox/index.asp?idpagina=575">RISCTOX</a>.<br/><br/>
En el supuesto de que no se rellenaran todos los campos o consideras que la información no coincide con la que tu dispones, puedes completarla y modificarla. Una vez que hayas completado la ficha pulsa sobre el botón GUARDAR PRODUCTO.</p>

<p><strong>* Resultados</strong>. Para hacer la evaluación preliminar del producto, pulsa sobre EVALUAR/COMPARAR. La información que verás es una evaluación básica sobre los niveles de riesgo de cada una de las sustancias que componen el producto sobre el que realizas la consulta.</p>

<p><strong>* Comparación</strong>. Para comparar varios productos elige de tu LISTA DE PRODUCTOS los que quieres comparar y pulsa el botón EVALUAR/COMPARAR. Los productos y sustancias se comparan por columnas, esto es  por tipos de riesgo (toxicidad aguda; ecotoxicidad; etc.). Además, se deben tener en cuenta las condiciones de uso del producto. A la vista de los niveles de riesgo identificados por la herramienta deberás de optar por el producto o sustancia que presente los niveles más bajos.</p>

<p align="center">
<input type=button class=boton value="Más información de como utilizarla" onclick='window.open("./dn_auto_mas_informacion_matriz.asp","Información","width=800,height=600,scrollbars=YES")'>
<br /><br />
</p>
<p align='center'>
<img src="imagenes/dn_ejemplo_evalua.jpg" alt="Ejemplo de tabla de evaluación" title="Ejemplo de tabla de evaluación"><br /><br />

<strong>SIEMPRE SE DEBE TOMAR EN CONSIDERACIÓN LA SITUACIÓN DE LA EMPRESA RESPECTO DE LAS CONDICIONES DE PREVENCIÓN Y GESTIÓN DEL RIESGO IMPLANTADAS.</strong></p>


<p class="titulo3">Interpretación de resultados</p>
<p>
Si el posible sustituto (sustancia/preparado) tiene una comparación final mejor que el producto actual en todas las columnas el problema de sustitución queda resuelto.
</p>
<p>
En la mayoría de los casos el resultado será tal que el posible sustituto tiene riesgo menor en algunas columnas y mayor en otras que el producto o sustancia a sustituir. Esto implica que habría que valorar los peligros potenciales, o en otras palabras, las columnas que tienen mayor peso en nuestra situación particular. Por ejemplo, si el proceso de producción implica grandes cantidades de residuos o subproductos, entonces el riesgo de toxicidad para el medio ambiente tendrá más énfasis. 
</p>
<p>
Si lo que queremos es una comparación en función de los riegos para la salud las dos primeras columnas (toxicidad) serán más relevantes. 
</p>
<p>
Cuando no existe información sobre ensayos de toxicidad o de sensibilización de la piel, el riesgo de toxicidad aguda se considera alto.
</p>
<p>
Cuando no existe información sobre ensayos de mutagenicidad, la sustancia o preparado debería categorizarse al menos en alto riesgo, en la columna de toxicidad crónica.
</p>
<p>
Si no existe información disponible de ensayos de efectos irritantes sobre la piel o mucosas, la sustancia o `preparado debería categorizarse al menos, en el apartado de bajo riesgo para toxicidad aguda.
</p>
<p>
Es importante recalcar que esta herramienta no tiene en cuenta las concentraciones de las sustancias. Cuando se valoran los productos en base a sus componentes, es posible que el nivel de riesgo sea más alto que el real, al no considerar las concentraciones de los mismos.

</p>


<p class="titulo3">Por qué impulsar la sustitución</p>
<p>La <a href="pdf/EcoSustanciasDefinitivaLEX.pdf">normativa básica</a> de referencia en riesgo químico establece como prioridad la eliminación del riesgo, por lo que la sustitución, en tanto que técnica preventiva, resulta una prioridad cuando no una obligación (cancerígenos, mutágenos y algunos tóxicos para la reproducción). Además, es prioritario eliminar o sustituir todas las sustancias que debido a su peligrosidad intrínseca presentan un nivel de riesgo inaceptable, incluidas en la <a href="http://www.istas.net/risctox/dn_risctox_negra.asp" target="_blank">lista negra de ISTAS</a>.</p>
<p>Al proponer a la empresa cualquier iniciativa de sustitución debemos concretar en primer lugar unos criterios para la búsqueda de alternativas. Después estableceremos unas etapas por las que avanzar en la materialización de la iniciativa. Puedes ampliar información en la <a href="http://www.istas.net/ecoinformas/web/abreenlace.asp?idenlace=2428">Guía para la sustitución de sustancias peligrosas</a>.</p>
<p align="center">
<%
if session("id_ecogente2")="" then
	'session("redirigir_tras_autentificar")="dn_auto_portada.asp"
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

<p>
  <!--
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
-->
  <br>
  Esta página ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundación de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a><br>
</p>
<center>©2012 www.istas.net, All rights reserved
</center>

				
		  </div>
				<p>&nbsp;</p>
			</div>
			
			
			<img src="../imagenes/pie_risctox.gif" width="708" border="0">

    	</div>
	<div id="sombra_abajo"></div>
</div>
<!--#include file="../../cookie_accept.asp" -->
</body>
</html>

<%
cerrarconexion
%>
