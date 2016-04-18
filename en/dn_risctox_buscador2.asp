<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_conexion.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->

<!--#include file="../dn_restringida.asp"-->

<%

'si busc est� vacio, mostramos formulario; si es 1, han dado a "buscar"; si es dos, han dado a paginaci�n

busc=request.form("busc")
busc = EliminaInyeccionSQL(busc)

	filtro=0 'este filtro diferencia a este buscador del de sustancias: si esta a true, muestra solo las que son toxicas y tienen alternativas
%>
	<!--#include file="dn_buscador_sustancias.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>RISCTOX: Toxic and hazardous substances database</title>
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

<link rel="stylesheet" type="text/css" href="../estructura.css">
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



	if ((frm.nombre.value.length < 3) && (frm.tipobus.options[frm.tipobus.selectedIndex].value == "parte"))

	{

		alert("Please type at least 3 characters to search by name");

	}

	else

	{

		frm.busc.value=1;
		frm.pag.value=1;
		frm.submit();

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

<td>

</td>

<td align="right"><input type="button" name="volver" class="boton" value="back to homepage" onClick="window.location='./index.asp';"></td>

</tr>

</table>

<p class=titulo3>RISCTOX: Toxic and hazardous substances database</p>


<form action="dn_risctox_buscador2.asp?busc=1" method="post" name="myform" onSubmit="primerapag();">
 <input type="hidden" name='busc' value='<%=busc%>' />
 <input type="hidden" name='pag' value='<%=pag%>' />
 <input type="hidden" name='hr' value='<%=hr%>' />
 <input type="hidden" name='arr' value='<%=arr%>' />
 <input type="hidden" name='ordenacion' value='<%=ordenacion%>' />
 <input type="hidden" name='nregs' value='<%=nregs%>' />
<table class="tabla3" width="90%" align="center">
<tr><td colspan="3" class="subtitulo3">Substance search</td></td></tr>
	<tr>
		<td align="right"><strong>Nombre</strong></td>
		<td><input type="text" name="nombre" value="<%=nombre%>" />
		<select name="tipobus">
		<option value="exacto" <%if tipobus="exacto" then response.write "selected"%>>exact name</option>

		<option value="parte" <%if tipobus="parte" then response.write "selected"%>>part of the name</option>

		</select></td>
	</tr>
	<tr>
		<td align="right"><strong>CAS/EC/Index No</strong></td>
		<td><input type="text" name="numero" value="<%=numero%>" /></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" value="Search" /> <input type="reset" value="Erase" /></td>
	</tr>
</table>

<%
if busc<>"" then
	if hr=0  then
%>
		<fieldset id="flashmsg"><legend class="advertencia"><strong>Warning</strong></legend>No records found that match your query.</fieldset>
<%
	else
		response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>" &hr& " records found. Showing from " &registroini+1& " to " &registrofin+1& ":</p>")



%>
		<%=tablares%>
<%
if hr>nregs then
%>
		<div align='center' style="margin:20px 10px; background-color: #3399CC; padding:3px;"><%paginacion%></div>
<%
end if
%>
<%
	end if
end if
%>
</form>


<!-- ############ FIN DE CONTENIDO ################## -->



<br>
<br>
This site has been developed by <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> - <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a>. This activity has been commisioned by <a href="http://www.etui.org/" target="_blank">ETUI</a> and supported by <a target="_blank" href="http://www.eeb.org/">EEB</a><br>

		  </div>
				<p>&nbsp;</p>
			</div>


			<img src="imagenes/pie_risctox.gif" width="708" border="0">


    		</div>
    	</div>
	<div id="sombra_abajo"></div>
</div>
</body>
</html>

<%
cerrarconexion
%>


