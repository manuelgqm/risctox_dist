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
<link rel="stylesheet" type="text/css" href="../estructura.css">
<link rel="stylesheet" type="text/css" href="../dn_estilos.css">

<script type="text/javascript" src="dn_auto_scripts.js"></script>

<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
window.onload=function(){
Nifty("div.dn_ncc_cuerpo","big"); 
}

</script>

</head>
<body>

<div id="dn_auto_frases_cuerpo" class="dn_ncc_cuerpo">
	<h2 class="dn_cabecera_2">&nbsp;&nbsp;&nbsp;Listado de frases R</h2>

	<form name="form_frases_r" action="dn_auto_frases_r2.asp" method="post">
	<p align="center">Seleccione las frases R que se aplican y pulse el botón "Guardar"</p>
	<center><input type="submit" class="boton2" name="guardar" value="Guardar"></center>
	<input type="hidden" name="idcampo" value="<%=EliminaInyeccionSQL(request.querystring("idcampo"))%>">

<%
sql = "SELECT id, frase, texto FROM dn_risc_frases_r ORDER BY frase"
set objRst = objConnection2.execute(sql)
%>
	<table border="0">
	<%
	do while (not objRst.eof)
	%>
		<tr>	
			<td class="frase_r" valign="top"><input type="checkbox" name="check" value="<%=objRst("frase")%>"></td>
			<td class="frase_r" valign="top"><strong><%=objRst("frase")%></strong> <%=objRst("texto")%></td>
		</tr>
	<%
		objRst.movenext
	loop
	%>
	</table>
	<center><input type="submit" class="boton2" name="guardar" value="Guardar"></center>
	</form>
	<br/><br/>
</div>

</body>
</html>

<%
cerrarconexion
%>
