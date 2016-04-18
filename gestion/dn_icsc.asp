<!--#include file="adovbs.inc"--><!--#include file="dn_conexion.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd"><html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1252"><title>Istas</title><link rel="stylesheet" type="text/css" href="dn_estilos.css"><link rel="stylesheet" type="text/css" href="dn_estilosmenu.css"><script type="text/javascript" src="niftycube.js"></script><script type="text/javascript">window.onload=function(){Nifty("ul#split h3","top");Nifty("ul#split div","bottom same-height");}</script></head><body><!--#include file="dn_menu.asp"--><h1>Importador de ICSC</h1>
<p>Este importador te permite pegar la relación entre números ICSC y números CAS.</p>
<p>Para ello, ve a <a href="http://www.mtas.es/insht/ipcsnspn/nspncas.htm" target="_blank">la página del ICSC</a>, copia el listado y pégalo en el siguiente campo. Pulsa el botón "Importar" y espera a la finalización del proceso.</p>
<p>En cada fila tiene que haber una única relación ICSC - CAS. Pega la relación completa (al importarse se borrará la relación antigua y se refrescará con el nuevo listado).</p>

<form name="icsc" action="dn_icsc2.asp" method="post">
<textarea name="listado" cols="80" rows="15"></textarea><br /><br />
<input type="submit" name="Importar" value="Importar" />
</form>

</body></html>

<% cerrarconexion %>
