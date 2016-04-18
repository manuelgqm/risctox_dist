<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<!--#include file="dn_funciones_comunes.asp"-->
<!--#include file="dn_funciones_texto.asp"-->

<!--#include file="dn_restringida.asp"-->

<%
'si busc está vacio, mostramos formulario; si es 1, han dado a "buscar"; si es dos, han dado a paginación 
busc=request.form("busc")
busc = EliminaInyeccionSQL(busc)
if busc="" then busc=1 'nada mas entrar, ya mostramos resultados
filtro="res" 
%>
	<!--#include file="dn_buscador_sustancias.asp"-->

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
                <td></td>
                <td align='right'><input type="button" name="volver" class="boton" value="Volver a la portada de risctox" onClick="window.location='./dn_risctox_buscador.asp';"></td>
                </tr>
                </table>
                
<p class=titulo3>RISCTOX: 	Sustancias que generan residuos peligrosos</p>


<form action="dn_risctox_res.asp?busc=1" method="post" name="myform" onSubmit="primerapag();">
 <input type="hidden" name='busc' value='<%=busc%>' />	
 <input type="hidden" name='pag' value='<%=pag%>' />	
 <input type="hidden" name='hr' value='<%=hr%>' />		
 <input type="hidden" name='arr' value='<%=arr%>' />
 <input type="hidden" name='ordenacion' value='<%=ordenacion%>' />
 <input type="hidden" name='nregs' value='<%=nregs%>' />				
<table class="tabla3" width="90%" align="center">
<tr>
  <td colspan="3" class="subtitulo3">Buscador</td></td></tr>
	<tr>
		<td align="right"><strong>Nombre</strong></td>
		<td><input type="text" name="nombre" value="<%=nombre%>" />
		<select name="tipobus">
		<option value="exacto" <%if tipobus="exacto" then response.write "selected"%>>nombre exacto</option>
		<option value="parte" <%if tipobus="parte" then response.write "selected"%>>parte del nombre</option>
		</select></td>
	</tr>
	<tr>
		<td align="right"><strong>Número CAS/CE/RD</strong></td>
		<td><input type="text" name="numero" value="<%=numero%>" /></td>
	</tr>	
	<tr>
		<td colspan="2" align="center"><input type="button" value="Buscar" class="boton" onClick="primerapag();" /> <input class="boton" type="reset" value="Borrar" /></td>
	</tr>
</table>

<%
if busc<>"" then
	if hr=0  then
%>
		<fieldset id="flashmsg"><legend class="advertencia"><strong>Advertencia</strong></legend>No se encontraron registros que coincidan con su consulta.</fieldset>
<%
	else
		response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>Se han encontrado " &hr& " registros. Se muestran registros del " &registroini+1& " al " &registrofin+1& ":</p>")
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
Esta página ha sido desarrollada por <a href="http://www.istas.ccoo.es/" target="_blank"><b>ISTAS</b></a> que es una Fundación de <a href="http://www.ccoo.es/" target="_blank"><font color="#FF0000"><b>CC.OO.</b></font></a><br>

				
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


