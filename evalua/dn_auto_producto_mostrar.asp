<!--#include file="../dn_conexion.asp"-->
<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_restringida.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: BBDD Evalúa lo que usas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="Dabne" />
<meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />

<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<link rel="stylesheet" type="text/css" href="estructura.css">

<script type="text/javascript" src="dn_auto_scripts.js"></script>

<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
window.onload=function(){
Nifty("div.dn_ncc_cuerpo","big"); 
}
</script>

</head>
<body>
<div id="contenedor">
	<div id="caja">
		<div class="texto">

<br/><br/>

<div id="dn_auto_producto_cuerpo" class="dn_ncc_cuerpo">

<%
id_producto = EliminaInyeccionSQL(request("id_producto"))

sql = "SELECT dn_auto_productos.nombre AS nombreproducto, dn_auto_procesos.nombre AS nombreproceso, dn_auto_productos.frases_r AS frases_r, dn_auto_estados.nombre AS nomestado, dn_auto_presiones.nombre AS nompresion, dn_auto_temperaturas.nombre AS nomtemperatura, dn_auto_inflamabilidades.nombre AS nominflamabilidad FROM dn_auto_productos INNER JOIN dn_auto_procesos ON dn_auto_productos.cod_proceso = dn_auto_procesos.cod INNER JOIN dn_auto_estados ON dn_auto_productos.cod_estado = dn_auto_estados.cod INNER JOIN dn_auto_presiones ON dn_auto_productos.cod_presion = dn_auto_presiones.cod INNER JOIN dn_auto_temperaturas ON dn_auto_productos.cod_temperatura = dn_auto_temperaturas.cod INNER JOIN dn_auto_inflamabilidades ON dn_auto_productos.cod_inflamabilidad = dn_auto_inflamabilidades.cod WHERE dn_auto_productos.id="&id_producto&" AND dn_auto_productos.id_ecogente="&session("id_ecogente2")
'response.write "<br>"&sql
set objRst = objConnection2.execute(sql)
if (objRst.eof) then
	' No se encontró el producto
%>
	<h2>Producto no encontrado</h2>
<%
else
	' Se encontró
%>
		<h2 class="dn_cabecera_2">&nbsp;&nbsp;&nbsp;Ficha de producto</h2>
		<table border="0">
			<tr>
				<th align="left">Nombre comercial</th>
				<th align="left">Frases R</th>	
			</tr>
			<tr>
				<td><%=objRst("nombreproducto")%></td>
				<td><%=objRst("frases_r")%></td>
			</tr>
			<tr>
				<th align="left" colspan="2">¿En qué tipo de proceso se emplea?</th>
			</tr>
			<tr>
				<td colspan="2"><%=objRst("nombreproceso")%></td>
			</tr>
			<tr>
				<th align="left">Estado físico</th>
				<th align="left">Presión de vapor</th>	
			</tr>
			<tr>
				<td><%=objRst("nomestado")%></td>
				<td><%=objRst("nompresion")%></td>
			</tr>
			<tr>
				<th align="left">Temperatura evaporación</th>
				<th align="left">Inflamabilidad</th>	
			</tr>
			<tr>
				<td><%=objRst("nomtemperatura")%></td>
				<td><%=objRst("nominflamabilidad")%></td>
			</tr>		
		</table>
<% 
	' Traemos los componentes
	sql2="SELECT dn_auto_componentes.nombre, dn_auto_componentes.numero_tipo, dn_auto_componentes.numero, dn_auto_componentes.frases_r,  porcentaje FROM dn_auto_componentes WHERE (dn_auto_componentes.id_producto = "&id_producto&")"

	'response.write "<br>"&sql2
	set objRst2=objConnection2.execute(sql2)

	if (not objRst2.eof) then
		nc = 0
		do while (not objRst2.eof)
			nc = nc + 1
%>
		<div id="comp<%=nc%>" class="componente">
				<p class="componente_titulo">Componente <%=nc%></p>

				<table border="0" width="50%">
					<tr>
						<th align="left">Tipo de número</th>
						<th align="left">Número identificativo</th>
					</tr>

					<tr>
						<td><%=ucase(objRst2("numero_tipo"))%></td>
						<td><%=objRst2("numero")%></td>
					</tr>
				</table>

				<table border="0" width="100%">
					<tr>
						<th align="left">Nombre</th>
					</tr>
					<tr>
						<td><%=objRst2("nombre")%></td>
					</tr>
				</table>

				<table border="0" width="100%">
					<tr>
						<th>Frases R</th>						
						<th>Concentración</th>
					</tr>
					<tr>
						<td><%=objRst2("frases_r")%></td>
						<td><%=objRst2("porcentaje")%> %</td>
					</tr>
				</table>				
		</div>

<%
			objRst2.movenext
		loop
		objRst2.close()
		set objRst2=nothing
	end if
end if

objRst.close()
set objRst=nothing
%>
</div>

<div align="center"><br/><input type="button" name="cerrar" value="Cerrar" onClick="window.close();" /><br/><br/></div>

 		</div>
	</div>
</div>
</body>
</html>

<%
cerrarconexion
%>
