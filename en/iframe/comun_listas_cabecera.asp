<!--#include file="../../adovbs.inc"-->
<!--#include file="../../dn_conexion.asp"-->
<!--#include file="../../dn_funciones_comunes.asp"-->
<!--#include file="../../dn_funciones_texto.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->

<!--#include file="../../dn_restringida.asp"-->

<%
'si busc está vacio, mostramos formulario; si es 1, han dado a "buscar"; si es dos, han dado a paginación
busc = request.form("busc")
busc = EliminaInyeccionSQL(busc)



if busc="" then busc=1 'nada mas entrar, ya mostramos resultados

%>

<!--#include file="dn_buscador_sustancias.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=titulo_lista%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="Risctox" />
<meta name="Author" content="SPL Sistemas de Información - www.spl-ssi.com" />
<meta name="description" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Subject" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Keywords" content="RISCTOX: Toxic and hazardous substances database" />
<meta name="Language" content="English" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />


<link rel="stylesheet" type="text/css" href="css/estructura.css">
<link rel="stylesheet" type="text/css" href="css/en.css">

<script type="text/javascript">
function cambiapag(paginadest)
{
	var frm = document.forms["myform"];
	frm.busc.value=2;
	frm.pag.value=paginadest;
	frm.submit();
}

function primerapag()
{
	var frm = document.forms["myform"];
	frm.busc.value=1;
	frm.pag.value=1;
	frm.submit();
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
                <td><a href="http://www.etuc.org/a/6023" target="_blank"><b>Trade Union priority list for REACH authorization</b></a></td>
                <td align='right'><input type="button" name="volver" class="boton" value="back" onClick="window.location='./dn_risctox_buscador.asp';"></td>
                </tr>
                </table>

<p class=titulo3><%=titulo_lista%></p>
