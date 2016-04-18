<!--#include file="../dn_conexion.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS Eval�a lo que usas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="Eval�a lo que usas" />
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
		<!--#include file="../dn_cabecera.asp"-->
		<div id="texto">
			
<div class="texto">
<img src=http://www.istas.net/recursos/IMG/ISTAS_01033.jpg align="right" hspace="20" vspace="10">

<p class=titulo3>�Qu� es Eval�a y Compara lo que Usas?</p>

<p>"Eval�a y Compara lo que Usas" es una herramienta para ayudar a valorar los riesgos para la salud y el medio ambiente de las sustancias y los productos qu�micos utilizados en los lugares de trabajo. Al permitir comparar los riesgos de varios productos, puede servir de ayuda en la b�squeda de alternativas que eviten o reduzcan el riesgo qu�mico en las empresas. 
</p>
<p>
Es un recurso basado en el <a href='http://www.hvbg.de/e/bia/pra/spalte/index.html' target='_blank'>Modelo de Columnas</a> desarrollado por el BIA (Instituto de Mutuas Profesionales Alemanas) y adaptado por el ISTAS.
</p>
<p>
NO ES UNA EVALUACI�N EN PROFUNDIDAD, sino tan s�lo una propuesta automatizada que realiza una valoraci�n preliminar del riesgo de un producto y sus componentes, y que depende del nivel de fiabilidad de la informaci�n que se facilite.
</p>
<p>
No se pretende con ella un dictamen, tan s�lo una orientaci�n, un primer paso, sencillo y pr�ctico, en el camino de la sustituci�n de sustancias peligrosas.
</p>
<p>
La matriz permite:
<ul>
	<li>Evaluar el riesgo que presentan los productos qu�micos que utilizas en tu empresa, a partir de la informaci�n existente en las Fichas de Datos de Seguridad (n�meros CAS y frases R).</li>
    <li>Comparar los riesgos que presentan diferentes productos, a fin de encontrar la alternativa m�s adecuada a las condiciones de uso que t� especifiques.</li>
</ul>
</p>
<p>
Incluye las siguientes variables: efectos agudos, efectos cr�nicos, ecotoxicidad, inflamabilidad y explosi�n, tipo de exposici�n y tipo de proceso de trabajo.
</p>
<p>
Permite clasificar cada una de las sustancias seg�n los siguientes niveles de riesgo: muy alto, alto, medio, bajo y muy bajo. 
</p>
<p>
En aquellas situaciones de riesgo en las que coincida m�s de una sustancia qu�mica, tendremos que realizar una evaluaci�n de la situaci�n de riesgo que resuma los resultados parciales de la evaluaci�n de cada una de las sustancias implicadas. 
</p>
<p>
Dado que en condiciones de multiexposici�n es probable que los efectos de cada una de las sustancias, se sumen o incluso multipliquen el resultado final, calificaremos le riesgo resultante de un producto, como m�nimo, igual al nivel de riesgo encontrado para alguna de las sustancias que lo componen.  
</p>
<p>
Esta matriz no considera las concentraciones en las que se encuentran las sustancias en los preparados, lo que puede dar lugar a una calificaci�n del riesgo de algunos productos superior a la establecida en las frases R que aparecen en su etiqueta o FDS.
</p>


<p align=center>
<%
if session("id_ecogente2")="" then
	'session("redirigir_tras_autentificar")="dn_auto_portada.asp"
%>
<input type=button class=boton value="Identif�quese para poder usar la herramienta" onclick=location.href="acceso.asp">
<%
else

%>
<input type=button class=boton value="Ir a Herramienta de Autoevaluaci�n" onclick=location.href="dn_auto_herramienta.asp">
<%
end if
%>
</p><br>

<p class="titulo3">�C�mo utilizarla?</p>
<p><strong>* Ficha cuestionario</strong>. Se rellena una ficha por cada producto que se quiera evaluar. La informaci�n necesaria se encuentra en las Fichas de Datos de Seguridad de los Productos (<a href="http://www.istas.net/risctox/index.asp?idpagina=529">FDS</a>). Los iconos <img src="imagenes/fam_istas/help.gif" align="absmiddle"> proporcionan informaci�n sobre c�mo completar cada apartado.<br/><br/>
Parte de la informaci�n sobre los componentes: nombre, n�meros de identificaci�n, Frases R se completa autom�ticamente si la sustancia se encuentra en la base de datos <a href="http://www.istas.net/risctox/index.asp?idpagina=575">RISCTOX</a>.<br/><br/>
En el supuesto de que no se rellenaran todos los campos o consideras que la informaci�n no coincide con la que tu dispones, puedes completarla y modificarla. Una vez que hayas completado la ficha pulsa sobre el bot�n GUARDAR PRODUCTO.</p>

<p><strong>* Resultados</strong>. Para hacer la evaluaci�n preliminar del producto, pulsa sobre EVALUAR/COMPARAR. La informaci�n que ver�s es una evaluaci�n b�sica sobre los niveles de riesgo de cada una de las sustancias que componen el producto sobre el que realizas la consulta.</p>

<p><strong>* Comparaci�n</strong>. Para comparar varios productos elige de tu LISTA DE PRODUCTOS los que quieres comparar y pulsa el bot�n EVALUAR/COMPARAR. Los productos y sustancias se comparan por columnas, esto es  por tipos de riesgo (toxicidad aguda; ecotoxicidad; etc.). Adem�s, se deben tener en cuenta las condiciones de uso del producto. A la vista de los niveles de riesgo identificados por la herramienta deber�s de optar por el producto o sustancia que presente los niveles m�s bajos.</p>

<p align="center">
<input type=button class=boton value="M�s informaci�n de como utilizarla" onclick='window.open("./dn_auto_mas_informacion_matriz.asp","Informaci�n","width=800,height=600,scrollbars=YES")'>
<br /><br />
</p>
<p align='center'>
<img src="imagenes/dn_ejemplo_evalua.jpg" alt="Ejemplo de tabla de evaluaci�n" title="Ejemplo de tabla de evaluaci�n"><br /><br />

<strong>SIEMPRE SE DEBE TOMAR EN CONSIDERACI�N LA SITUACI�N DE LA EMPRESA RESPECTO DE LAS CONDICIONES DE PREVENCI�N Y GESTI�N DEL RIESGO IMPLANTADAS.</strong></p>


<p class="titulo3">Interpretaci�n de resultados</p>
<p>
Si el posible sustituto (sustancia/preparado) tiene una comparaci�n final mejor que el producto actual en todas las columnas el problema de sustituci�n queda resuelto.
</p>
<p>
En la mayor�a de los casos el resultado ser� tal que el posible sustituto tiene riesgo menor en algunas columnas y mayor en otras que el producto o sustancia a sustituir. Esto implica que habr�a que valorar los peligros potenciales, o en otras palabras, las columnas que tienen mayor peso en nuestra situaci�n particular. Por ejemplo, si el proceso de producci�n implica grandes cantidades de residuos o subproductos, entonces el riesgo de toxicidad para el medio ambiente tendr� m�s �nfasis. 
</p>
<p>
Si lo que queremos es una comparaci�n en funci�n de los riegos para la salud las dos primeras columnas (toxicidad) ser�n m�s relevantes. 
</p>
<p>
Cuando no existe informaci�n sobre ensayos de toxicidad o de sensibilizaci�n de la piel, el riesgo de toxicidad aguda se considera alto.
</p>
<p>
Cuando no existe informaci�n sobre ensayos de mutagenicidad, la sustancia o preparado deber�a categorizarse al menos en alto riesgo, en la columna de toxicidad cr�nica.
</p>
<p>
Si no existe informaci�n disponible de ensayos de efectos irritantes sobre la piel o mucosas, la sustancia o `preparado deber�a categorizarse al menos, en el apartado de bajo riesgo para toxicidad aguda.
</p>
<p>
Es importante recalcar que esta herramienta no tiene en cuenta las concentraciones de las sustancias. Cuando se valoran los productos en base a sus componentes, es posible que el nivel de riesgo sea m�s alto que el real, al no considerar las concentraciones de los mismos.

</p>


<p class="titulo3">Por qu� impulsar la sustituci�n</p>
<p>La <a href="pdf/EcoSustanciasDefinitivaLEX.pdf">normativa b�sica</a> de referencia en riesgo qu�mico establece como prioridad la eliminaci�n del riesgo, por lo que la sustituci�n, en tanto que t�cnica preventiva, resulta una prioridad cuando no una obligaci�n (cancer�genos, mut�genos y algunos t�xicos para la reproducci�n). Adem�s, es prioritario eliminar o sustituir todas las sustancias que debido a su peligrosidad intr�nseca presentan un nivel de riesgo inaceptable, incluidas en la <a href="http://www.istas.net/risctox/dn_risctox_negra.asp" target="_blank">lista negra de ISTAS</a>.</p>
<p>Al proponer a la empresa cualquier iniciativa de sustituci�n debemos concretar en primer lugar unos criterios para la b�squeda de alternativas. Despu�s estableceremos unas etapas por las que avanzar en la materializaci�n de la iniciativa. Puedes ampliar informaci�n en la <a href="http://www.istas.net/ecoinformas/web/abreenlace.asp?idenlace=2428">Gu�a para la sustituci�n de sustancias peligrosas</a>.</p>
<p align="center">
<%
if session("id_ecogente2")="" then
	'session("redirigir_tras_autentificar")="dn_auto_portada.asp"
%>
<input type=button class=boton value="Identif�quese para poder usar la herramienta" onclick=location.href="acceso.asp">
<%
else
%>
<input type=button class=boton value="Ir a Herramienta de Autoevaluaci�n" onclick=location.href="dn_auto_herramienta.asp">
<%
end if
%>
</p>

<p>
  <!--
<a id="adaptado_istas"></a>
<p class="titulo3">Notas sobre la adaptaci�n de ISTAS</p>
<p>A diferencia del modelo de columnas, esta herramienta considera sustancias de muy alto riesgo de toxicidad cr�nica:</p>

<p>
	<ul>
		<li>Las sustancias cancer�genas C3 (R40) y mut�genas M3 (R68) seg�n el RD 363/1995 y las sustancias cancer�genas 1, 2A y 2B seg�n IARC.</li>
		<li>Las sustancias t�xicas para la reproducci�n: R60, R61, R62 y R63.</li>
		<li>Las sustancias bioacumulables (R33) y que se acumulan en la leche materna (R64).</li>
		<li>Las sustancias sensibilizantes, neurot�xicas y disruptores endocrinos.</li>
	</ul>
</p>

<p>Adem�s, considera de muy alto riesgo para el medio ambiente las sustancias t�xicas, persistentes y bioacumulativas y los disruptores endocrinos.</p>

<p>La herramienta utiliza los listados de sustancias peligrosas de la base de datos RISCTOX elaborada por ISTAS.
<br/><br/></p>
-->
  <br>
  Esta p�gina ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundaci�n de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a><br>
</p>
<center>�2012 www.istas.net, All rights reserved
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
