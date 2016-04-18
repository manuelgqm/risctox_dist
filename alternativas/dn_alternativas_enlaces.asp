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

<p class=titulo3>Enlaces - Fuentes de información generales</p>

<%
	sql = "select titulo, enlace, texto from dn_alter_enlaces where clasificacion=1"
	set objRst = objConnection2.execute(sql)
	contador=1
	while not objRst.eof 
		response.write "<h3><a class='enlpleg' style='cursor:pointer' id='e"&contador&"' onclick='cambia("&contador&")'>"&objrst("titulo")&"</a></h3>"
		response.write "<p class='txpleg' id='t"&contador&"'>"&objrst("texto")&"<br /><br />"
		response.write "<a href='"&objrst("enlace")&"' target='_blank'>"&objrst("enlace")&"</a></p>"
		contador = contador + 1
		objrst.movenext()
	wend
%>



<p class=titulo3><br />
  Enlaces - 
      Informaci&oacute;n sobre eliminaci&oacute;n/sustituci&oacute;n<br />
    &nbsp;</p>
	
 <%
	sql = "select titulo, enlace, texto from dn_alter_enlaces where clasificacion=2"
	set objRst = objConnection2.execute(sql)
	while not objRst.eof 
		response.write "<h3><a class='enlpleg' id='e"&contador&"' onclick='cambia("&contador&")'>"&objrst("titulo")&"</a></h3>"
		response.write "<p class='txpleg' id='t"&contador&"'>"&objrst("texto")&"<br /><br />"
		response.write "<a href='"&objrst("enlace")&"' target='_blank'>"&objrst("enlace")&"</a></p>"
		contador = contador + 1
		objrst.movenext()
	wend
%>  


<!-- ############ FIN DE CONTENIDO ################## -->
<!--#include file="spl_pie.inc.asp"-->

<%
cerrarconexion
%>
