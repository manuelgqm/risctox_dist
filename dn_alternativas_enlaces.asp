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
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<style type="text/css">
<!--
/*titulos*/
#texto h3 {  margin:20px 0 10px 0; font-size:14px; font-weight:bold; padding:0; }
.enlpleg { background : #fff url(imagenes/dn_desplegar.gif)  top left no-repeat;  }
.enldespleg { background : #fff url(imagenes/dn_plegar.gif)  top left no-repeat;  }
.enlpleg, .enldespleg {color: #00ac5a; padding-left:25px; text-decoration:none; }
/*descripciones: por defecto, plegadas*/
.txpleg { display:none  }
.txdespleg { display:block  }
-->
</style>
<script type="text/javascript">
function cambia(numero)
{
	var cabecera = document.getElementById("e"+numero);

	var texto = document.getElementById("t"+numero);



	if (cabecera.className == "enlpleg")

	{

		cabecera.className = "enldespleg";

		texto.className = "txdespleg";

	}

	else

	{

		cabecera.className = "enlpleg";

		texto.className = "txpleg";

	}

}
</script>
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
<td><p class=campo>Est&aacute;s en: <a href=index.asp?idpagina=550>Plataforma prevención de riesgo químico</a>&nbsp;&gt;&nbsp;<a href="dn_alternativas_portada.asp">BBDD Alternativas</a> &gt; Enlaces </p></td>
<td><input type="button" name="volver" class="boton2" value="Volver a la portada de Alternativas" onclick="window.location='dn_alternativas_portada.asp';"></td>
</tr>
</table>

<p class=titulo3>Enlaces - Fuentes de información generales</p>

      <h3><a class="enlpleg" id="e1" onclick="cambia(1)">INSHT</a></h3>
    <p class="txpleg" id="t1">El Instituto Nacional de Salud e Higiene en el trabajo, contiene un sinf&iacute;n de publicaciones sobre legislaci&oacute;n y normalizaci&oacute;n, gu&iacute;as t&eacute;cnicas, gu&iacute;as pr&aacute;cticas y metodolog&iacute;as para la acci&oacute;n preventiva, estudios, manuales, protocolos, etc.<br /><br />
      <a href="http://www.mtas.es/insht/information/lib_tot.htm" target="_blank">http://www.mtas.es/insht/information/lib_tot.htm</a></p>
    
      <h3><a class="enlpleg" id="e2" onclick="cambia(2)">Agencia Europea para la seguridad y la salud Laboral</a></h3>
    <p class="txpleg" id="t2"><a href="http://osha.europa.eu/OSHA" target="_blank">http://osha.europa.eu/OSHA</a> </p>
	
  
      <h3><a class="enlpleg" id="e3" onclick="cambia(3)">IHOBE.S.A. </a></h3>
    <p class="txpleg" id="t3">La p&aacute;gina web de la Sociedad P&uacute;blica de gesti&oacute;n Ambiental del Gobierno vasco incluye copias electr&oacute;nicas de varias de sus publicaciones sobre producci&oacute;n limpia y minimizaci&oacute;n de residuos en diferentes sectores industriales. Tambi&eacute;n incluye enlaces interesantes a asociaciones empresariales sectoriales e institutos tecnol&oacute;gicos.<br /><br />
      <a href="http://www.ihobe.net/" target="_blank">http://www.ihobe.net/</a> </p>
	
    <h3><a class="enlpleg" id="e4" onclick="cambia(4)">CEMA: ejemplos de actuaciones de minimizaci&oacute;n de residuos y emisiones</a></h3>
    <p class="txpleg" id="t4">El CEMA es tambi&eacute;n el Centro de Actividades Regionales para la<br />
      Producci&oacute;n Limpia encargado de promover y difundir la prevenci&oacute;n de la ontaminaci&oacute;n en los pa&iacute;ses de la cuenca mediterr&aacute;nea. En su p&aacute;gina web se pueden encontrar ejemplos de sustituci&oacute;n de disolventes en empresas.<br />
      Una de las iniciativas de difusi&oacute;n es la publicaci&oacute;n de las fichas de experiencias MEDCLEAN .<br /><br />
      <a href="http://www.cprac.org/index_cast.htm" target="_blank">http://www.cprac.org/index_cast.htm</a></p>	  
	  
      <h3><a class="enlpleg" id="e5" onclick="cambia(5)">ISTAS </a></h3>
    <p class="txpleg" id="t5"> Base de datos de alternativas: ofrece documentos sobre productos, procesos y tecnolog&iacute;as alternativas que pueden ayudar a prevenir el riesgo qu&iacute;mico en tu empresa. La b&uacute;squeda se puede realizar por productos, procesos o sectores. Se puede consultar la lista completa de documentos incluidos en la base de datos y tambi&eacute;n se puede acceder a las alternativas desde la base de datos de sustancias RISCTOX. Se han incluido documentos en espa&ntilde;ol, catal&aacute;n, eusquera y gallego.<br /><br />
      <a href="http://www.istas.net/ecoinformas" target="_blank">http://www.istas.net/ecoinformas</a></p>
   
     <h3><a class="enlpleg" id="e6" onclick="cambia(6)"> EPA Portal Espa&ntilde;ol de la Agencia de Protecci&oacute;n Ambiental de los Estados Unidos. </a></h3>
    <p class="txpleg" id="t6">Contiene informaci&oacute;n &uacute;til en espa&ntilde;ol y ocasionalmente en ingl&eacute;s referente a la protecci&oacute;n de la salud humana y los recursos naturales, con secciones sobre prevenci&oacute;n de la contaminaci&oacute;n.<br /><br />Podemos encontrar informaci&oacute;n interesante sobre alternativas para prevenir la contaminaci&oacute;n en diferentes procesos y sectores, as&iacute; como gu&iacute;as para la reducci&oacute;n de la exposici&oacute;n de sustancias peligrosas a los trabajadores, buenas pr&aacute;cticas, etc.<br /><br />
      <a href="http://www.epa.gov/espanol/" target="_blank">http://www.epa.gov/espanol/</a></p>
	
    <h3><a class="enlpleg" id="e7" onclick="cambia(7)">PAM: Plan De Acci&oacute;n Para El Mediterr&aacute;neo</a></h3>
    <p class="txpleg" id="t7">El Plan de Acci&oacute;n para la protecci&oacute;n y el desarrollo de la cuenca del Mediterr&aacute;neo (PAM) (<a href="http://www.unepmap.org/homespa.asp" target="_blank">http://www.unepmap.org/homespa.asp</a> ) pertenece al Programa de las Naciones Unidas para el Medio Ambiente (PNUMA) y tiene como objetivo la protecci&oacute;n y mejora del medio ambiente y el desarrollo de la Regi&oacute;n, bas&aacute;ndose en los principios de la sostenibilidad.<br />
      Veinti&uacute;n pa&iacute;ses mediterr&aacute;neos, junto con la Uni&oacute;n Europea, est&aacute;n implicados en el Plan de Acci&oacute;n del Mediterr&aacute;neo, para la cooperaci&oacute;n y el desarrollo de la Regi&oacute;n.<br /><br />A trav&eacute;s del PAM las Partes Contratantes del Convenio de Barcelona y sus Protocolos quieren hacer frente a los desaf&iacute;os de la protecci&oacute;n del medio ambiente marino y de las costas del &aacute;rea del Mediterr&aacute;neo, impulsando programas, a nivel regional y nacional, para alcanzar el desarrollo sostenible en el &aacute;rea, para lo cual, la p&aacute;gina web contiene diversos documentos e informes t&eacute;cnicos, que promocionan el desarrollo sostenible y tratan de reducir la contaminaci&oacute;n.</p>


    <h3><a class="enlpleg" id="e8" onclick="cambia(8)">Productosostenible.net </a></h3>
    <p class="txpleg" id="t8"> Es un portal con informaci&oacute;n para la mejora ambiental de los productos industriales, dirigido a todos los agentes que intervienen a lo largo del Ciclo de Vida de un producto:<br />
      <br />
      Dise&ntilde;adores y fabricantes de productos<br />
      Administraci&oacute;n y Ciudadanos, como consumidores finales de los productos<br />
      Universidades y centros de ense&ntilde;anza<br />
      Los contenidos disponible surgen del trabajo de investigaci&oacute;n y recopilaci&oacute;n de informaci&oacute;n que se realiza desde los centros de documentaci&oacute;n de las Aulas de Ecodise&ntilde;o de la Universidad del Pa&iacute;s Vasco y de Universidad de Mondrag&oacute;n. En este trabajo de recopilaci&oacute;n participan adem&aacute;s diferentes empresas, asociaciones y centros tecnol&oacute;gicos. <br />
      La web, tiene contenidos tanto de PRODUCTOS, con una informaci&oacute;n detallada de sus caracter&iacute;sticas ambientales, como SECTORES, con informaci&oacute;n t&eacute;cnica de inter&eacute;s para la mejora de los productos de un determinado sector:<br />
      Definici&oacute;n de la problem&aacute;tica ambiental de los productos del sector<br />
      Herramientas de mejora y servicios de apoyo para las empresas<br />
      Legislaci&oacute;n ambiental de aplicaci&oacute;n a los productos<br />
      Estrategias de mejora ambiental<br />
      Sistemas de reconocimiento ambiental de productos<br />
      Casos pr&aacute;cticos<br />
      <br />
      <a href="http://www.productosostenible.net/pags/AP/Ap_inicio/index.asp?cod=449"  target="_blank">http://www.productosostenible.net/pags/AP/Ap_inicio/index.asp?cod=449</a></p>
    <h3><a class="enlpleg" id="e9" onclick="cambia(9)">CDC </a></h3>
    <p class="txpleg" id="t9"> El centro de los CDC encargado del tema de la salud ocupacional es el<a href="http://www.cdc.gov/spanish/niosh/index.html" target="_blank"> Instituto Nacional para la Seguridad y Salud Ocupacional (NIOSH).</a></p>
	
    <h3><a class="enlpleg" id="e10" onclick="cambia(10)">MedLine PLus</a></h3>
    <p class="txpleg" id="t10">MedlinePlus le ayuda encontrar las respuestas que usted busca en temas de salud. MedlinePlus ha recopilado la informaci&oacute;n m&aacute;s confiable proveniente de fuentes autorizadas tales como la Biblioteca Nacional de Medicina y los Institutos Nacionales de la Salud, as&iacute; como otras agencias gubernamentales y organizaciones de servicios para la salud. MedlinePlus tambi&eacute;n le ofrece mucha informaci&oacute;n sobre medicamentos, una enciclopedia m&eacute;dica ilustrada, programas interactivos para el paciente y las m&aacute;s recientes noticias acerca de la salud.<br />
      <br />
      <a href="http://www.nlm.nih.gov/medlineplus/spanish/occupationalhealth.html"  target="_blank">http://www.nlm.nih.gov/medlineplus/spanish/occupationalhealth.html</a><br />
     
	 <h3><a class="enlpleg" id="e11" onclick="cambia(11)"> OSHA-Am&eacute;rica</a></h3>
    <p class="txpleg" id="t11"> Es una agencia que pretende asegurar la Seguridad y Salud ocupacionales en Am&eacute;rica, estableciendo y haciendo cumplir normas, ofrecimiento de adiestramientos y educaci&oacute;n, estableciendo asociaciones y motivando a un mejoramiento continuo en la seguridad y salud en el lugar de trabajo. Tienen tambi&eacute;n programas de asesoramiento y de asistencia t&eacute;cnica.<br />
      Posee un portal en espa&ntilde;ol;<a href="http://www.osha.gov/as/opa/spanish/"  target="_blank"> http://www.osha.gov/as/opa/spanish/</a><br />
      Donde en el apartado, eTools en Espa&ntilde;ol , podr&aacute;s encontrar experiencias, publicaciones, documentos, etc, sobre Alternativas y buenas pr&aacute;cticas en diversos sectores.</p>
   
    <h3><a class="enlpleg" id="e12" onclick="cambia(12)">OSHA-Europa</a></h3>
    <p class="txpleg" id="t12">La agencia Europea para la seguridad y salud en el trabajo (<a href="http://osha.europa.eu/info"  target="_blank">http://osha.europa.eu/info</a>) pretende enfrentarse a la diversa problem&aacute;tica que entra&ntilde;a la seguridad y la salud en el trabajo (SST) y a la necesidad de incrementar la sensibilizaci&oacute;n en el centro de trabajo como una tarea que rebasa las capacidades y competencias de un solo Estado miembro. De ah&iacute; que en 1996 se crease la Agencia Europea para la Seguridad y la Salud en el Trabajo: para recopilar, analizar y promover informaci&oacute;n relacionada con la SST. La misi&oacute;n de la Agencia es hacer los puestos de trabajo europeos m&aacute;s sanos, seguros y productivos, y en particular fomentar una cultura de la prevenci&oacute;n en el lugar de trabajo.<br />
      Para ello, podemos encontrar buenas pr&aacute;cticas, documentos y campa&ntilde;as de prevenci&oacute;n y promoci&oacute;n de la salud.</p>
 
   <h3><a class="enlpleg" id="e13" onclick="cambia(13)">Greenpeace Espa&ntilde;a</a></h3>
    <p class="txpleg" id="t13">Portal donde se ofrece una completa informaci&oacute;n acerca de sustancias qu&iacute;micas peligrosas para la salud y el medio ambiente, tales como PVC&acute;s, COP&acute;s, DDT, Residuos, etc.<br />
      Adem&aacute;s puedes encontrar proyectos, informes, campa&ntilde;as de sensibilizaci&oacute;n y lucha contra los contaminantes, as&iacute; como una extensa recopilaci&oacute;n de casos donde la sustituci&oacute;n de productos t&oacute;xicos ha sido posible.<br />
      <br />
      <a href="http://www.greenpeace.org/espana/campaigns/t-xicos"  target="_blank">http://www.greenpeace.org/espana/campaigns/t-xicos</a></p>

   <h3><a class="enlpleg" id="e14" onclick="cambia(14)">WWF/ Adena</a></h3>
    <p class="txpleg" id="t14">WWF/Adena que ha emprendido una gran campa&ntilde;a en Europa a trav&eacute;s de su campa&ntilde;a DetoX , cuyo objetivo principal es asegurarse que la nueva normativa europea sobre el registro, la evaluaci&oacute;n y la autorizaci&oacute;n de sustancias qu&iacute;micas producidas e importadas en Europa (REACH seg&uacute;n sus siglas en ingl&eacute;s) sea bastante &ldquo;fuerte&rdquo; para eliminar aquellas sustancias qu&iacute;micas que representan un serio peligro para el medio ambiente y la salud.<br />
      Tambi&eacute;n ofrece en su p&aacute;gina informaci&oacute;n sobre los peligros de sustancias qu&iacute;micas, as&iacute; como consejos pr&aacute;cticos para reducir el riesgo de exposici&oacute;n a sustancias t&oacute;xicas.<br />
      <br />
      <a href="http://www.wwf.es/toxicos/toxicos.php"  target="_blank">http://www.wwf.es/toxicos/toxicos.php</a></p>
  
    <h3><a class="enlpleg" id="e15" onclick="cambia(15)">RAAA: Red de Acci&oacute;n en Agricultura Alternativa</a></h3>
    <p class="txpleg" id="t15">La Red de Acci&oacute;n en Agricultura Alternativa,es un movimiento que agrupa instituciones y personalidades del sector agrario en todo el Per&uacute;, cuya perspectiva es contribuir al desarrollo de la agricultura sostenible y la preservaci&oacute;n del ambiente. Forman parte de redes internacionales como RAP-AL y PAN Internacional.<br />
      Tiene informaci&oacute;n sobre agricultura org&aacute;nica, biotecnolog&iacute;a y otras alternativas al uso de pesticidas.<br />
      <br />
      <a href="http://www.raaa.org/"  target="_blank">http://www.raaa.org/</a>
    </p>
  
   <h3><a class="enlpleg" id="e16" onclick="cambia(16)">ATSDR</a></h3>
    <p class="txpleg" id="t16">Es la Agencia para sustancias t&oacute;xicas y el registro de enfermedades, perteneciente al Departamento de Salud y Servicios Humanos de EE.UU. [ En ingl&eacute;s. ], y su f&iacute;n es servir al p&uacute;blico usando la mejor ciencia, tomando acciones de salud p&uacute;blica que corresponden y proporcionar informaci&oacute;n de salud confiable, para prevenir exposiciones nocivas y enfermedades relacionadas a sustancias t&oacute;xicas.<br />
      La ATSDR tiene que, por Mandato del Congreso, realizar una serie de funciones relacionadas a los efectos sobre la salud humana de sustancias peligrosas que se encuentran en el medio ambiente. Estas funciones incluyen evaluaciones de salud p&uacute;blica de sitios que contienen desperdicios, consultas de salud con respecto a sustancias peligrosas en espec&iacute;fico, vigilancias y registros de salud, respuestas a emergencias debido a la emisi&oacute;n o derrame imprevisto de sustancias peligrosas, investigaci&oacute;n aplicada en respaldo a evaluaciones de salud p&uacute;blica, elaboraci&oacute;n y difusi&oacute;n de informaci&oacute;n y educaci&oacute;n y adiestramiento relacionados a sustancias peligrosas.<br />
      <br />
      <a href="http://www.atsdr.cdc.gov/es/es_about.html"  target="_blank">http://www.atsdr.cdc.gov/es/es_about.html</a></p>
  
   <h3><a class="enlpleg" id="e17" onclick="cambia(17)">Curso sobre prevenci&oacute;n, preparaci&oacute;n y respuesta a los desastres producidos por productos qu&iacute;micos</a></h3><br clear="all" />
    <p class="txpleg" id="t17">En esta p&aacute;gina web, se tratan los documentos t&eacute;cnicos del curso internacional &quot;Prevenci&oacute;n, preparaci&oacute;n y respuesta a desastres por productos qu&iacute;micos peligrosos&quot;, organizado por CETESB, la Compa&ntilde;&iacute;a de Tecnolog&iacute;a de Saneamiento Ambiental del estado de Sao Paulo (Brasil), en estrecha colaboraci&oacute;n con la Organizaci&oacute;n Panamericana de la Salud (OPS/OMS). Es un material preparado por especialistas de distintos pa&iacute;ses de la Am&eacute;rica Latina y el Caribe, que hace un recorrido por los m&aacute;s importantes aspectos te&oacute;ricos y pr&aacute;cticos sobre la prevenci&oacute;n y la respuesta a los desastres producidos por sustancias qu&iacute;micas. La primera edici&oacute;n de este curso se celebr&oacute; en San Paulo (Brasil) en octubre de 1999. <br />
      Entre los documentos del curso, pueden encontrarse conceptos generales vinculados a los accidentes qu&iacute;micos y clasificaci&oacute;n de materiales peligrosos y toxicolog&iacute;a (ponencias 1 a 4); planificaci&oacute;n y acciones de respuesta general y organizaci&oacute;n institucional requerida para llevarlas a cabo (5 a 10); acciones de respuesta espec&iacute;ficamente m&eacute;dica destinada a las v&iacute;ctimas de este tipo de desastres (11 a 14); descripci&oacute;n de equipos de protecci&oacute;n y de aparatos de monitoreo del medioambiente (15 y 16), y finalmente la descripci&oacute;n de algunas experiencias puntuales de organizaci&oacute;n para la prevenci&oacute;n y respuesta en caso de desastres qu&iacute;micos (17 a 19). <br />
      CETESB fue declarado Centro Colaborador de la OMS por su experiencia y capacitaci&oacute;n en la prevenci&oacute;n y respuesta a accidentes con sustancias qu&iacute;micas. Por parte de la OPS, han intervenido en la organizaci&oacute;n del curso el Programa de Preparativos para Casos de Desastres, y el Centro Panamericano de Ingenier&iacute;a Sanitaria y Ciencias del Ambiente (CEPIS).<br /><br />
      <a href="http://www.disaster-info.net/quimicos/index.htm"  target="_blank">http://www.disaster-info.net/quimicos/index.htm</a></p>

<p class=titulo3><br />
  Enlaces - 
      Informaci&oacute;n sobre eliminaci&oacute;n/sustituci&oacute;n<br />
    &nbsp;</p>
	
   <h3><a class="enlpleg" id="e18" onclick="cambia(18)">CLEANTOOL (Limpieza de metales)</a></h3>
    <p class="txpleg" id="t18">CLEANTOOL ofrece una base de datos de alternativas y de buenas pr&aacute;cticas en los procesos de limpieza de superficies met&aacute;licas. Facilita informaci&oacute;n para ayudar a las empresas elegir el proceso de limpieza &oacute;ptimo seg&uacute;n las necesidades de su actividad.<br />
      <br />
      <a href="http://www.cleantool.org/lang/sp/start_sp.htm"  target="_blank">http://www.cleantool.org/lang/sp/start_sp.htm</a></p>
    <h3><a class="enlpleg" id="e19" onclick="cambia(19)">PPGEMS</a></h3>
    <p class="txpleg" id="t19">Portal del Toxic Use Reduction Institute de Massachussets que ofrece cientos de enlaces a p&aacute;ginas sobre prevenci&oacute;n de la contaminaci&oacute;n, organizadas por productos o industrias, sustancias qu&iacute;micas o residuos, herramientas de gesti&oacute;n y procesos.<br /><br />
      <a href="http://www.p2gems.org/"  target="_blank">http://www.p2gems.org/</a></p>
    
     <h3><a class="enlpleg" id="e20" onclick="cambia(20)"> PAN Alternativas a pesticidas</a></h3>
    <p class="txpleg" id="t20">Este sitio contiene informaci&oacute;n sobre alternativas al uso de plaguicidas, clasificadas por plaga y por cultivo vegetal.<br /><br />
      <a href="http://www.pesticideinfo.org/Index.html" target="_blank">http://www.pesticideinfo.org/Index.html</a></p>
	
      <h3><a class="enlpleg" id="e21" onclick="cambia(21)">INSHT</a></h3>
    <p class="txpleg" id="t21">NTP 673: La sustituci&oacute;n de agentes qu&iacute;micos peligrosos: aspectos generales.<br />
      NTP 712: Sustituci&oacute;n de agentes qu&iacute;micos peligrosos (II): criterios y modelos pr&aacute;cticos.<br /><br />
      <a href="http://www.mtas.es/insht/ntp" target="_blank">http://www.mtas.es/insht/ntp</a></p>
  
      
	  <h3><a class="enlpleg" id="e22" onclick="cambia(22)">SAGE Data Base </a></h3>
    <p class="txpleg" id="t22"> Alternativas para la sustituci&oacute;n de disolventes industriales<br /><br />
      <a href="http://sage.rti.org/" target="_blank">http://sage.rti.org/</a></p>
    
	<h3><a class="enlpleg" id="e23" onclick="cambia(23)">CAGE Data Base </a></h3>
    <p class="txpleg" id="t23"> Alternativas para la sustituci&oacute;n de pinturas y revestimientos industriales.<br /><br />
      <a href="http://cage.rti.org" target="_blank">http://cage.rti.org</a></p>


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


