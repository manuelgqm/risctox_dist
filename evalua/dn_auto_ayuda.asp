<!--#include file="../dn_conexion.asp"-->
<!--#include file="../adovbs.inc"-->
<!--#include file="../dn_restringida.asp"-->
<!--#include file="../dn_funciones_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ISTAS: BBDD Evalúa lo que usas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-15" />
<meta name="Title" content="ECOinformas" />
<meta name="Author" content="XiP multimèdia" />
<meta name="description" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Subject" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Keywords" content="Información, formación y asesoramiento sobre medio ambiente para trabajadores de PYME" />
<meta name="Language" content="Spanish" />
<meta name="Revisit" content="15 days" />
<meta name="Distribution" content="Global" />
<meta name="Robots" content="All" />

<link rel="stylesheet" type="text/css" href="../dn_estilos.css">
<link rel="stylesheet" type="text/css" href="../estructura.css">

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
razon = EliminaInyeccionSQL(quitartildes(request("razon")))
'response.write "["&razon&"]"

if es_frase_r(razon) then
	sql = "SELECT texto as explicacion FROM dn_risc_frases_r WHERE frase='" &razon& "'"
else
	sql = "SELECT explicacion FROM dn_auto_ayuda WHERE razon='" &razon& "'"
end if
'response.write "<br>"&sql
set objRst = objConnection2.execute(sql)
if (objRst.eof) then
	' No se encontró ayuda
%>
	<h2 class="dn_cabecera_2">&nbsp;&nbsp;&nbsp;<%=razon%></h2>
<%
else
	' Se encontró
%>
		<h2 class="dn_cabecera_2">&nbsp;&nbsp;&nbsp;<%=razon%></h2>
		<table border="0">
			<tr>
				<td><%=objRst("explicacion")%></td>
			</tr>
		</table>
<% 
end if

objRst.close()
set objRst=nothing
%>
</div>

<p align="center"><input type="button" name="cerrar" onClick="window.close();" value="cerrar ventana" class="boton2"></p>

 		</div>
	</div>
</div>
</body>
</html>

<%
cerrarconexion
%>
