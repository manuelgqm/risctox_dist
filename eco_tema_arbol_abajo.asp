<!--#include file="eco_conexion3.asp"-->
<html>
<head>
<title>Espacio web ISTAS:::buscador de páginas</title>
<link rel="stylesheet" type="text/css" href="panelcontrol.css">
<base target="contenido">
</head>

<body bgcolor="#FEFEFE" topmargin="5" leftmargin="5">
<form name="formulario" action="eco_buscarpaginas.asp">
<table>
<tr><td class="cue_celda"><b>Buscar páginas</b></td></tr>
<tr><td class="cue_fuente">número:&nbsp;<input type="text" size="25" class="campo" name="numero"></td></tr>
<tr><td class="cue_fuente">texto:&nbsp;<input type="text" size="25" class="campo" name="buscar"></td></tr>
<tr><td class="cue_fuente">fechas:&nbsp;<input type="text" size="10" class="campo" name="fecini">&nbsp;y&nbsp;<input type="text" size="10" class="campo" name="fecfin"></td></tr>
<tr><td class="cue_fuente" align="center"><input type="submit" value="BUSCAR" class="boton">&nbsp;&nbsp;<input type="button" value="VACIAR PAPELERA" class="boton" onclick="parent.parent.frames.contenido.location='vaciar_papelera.asp'"></td></tr>
</table>
</form>
</body>
</html>