<!--#include file="../EliminaInyeccionSQL.asp"-->
<%
	session( "idgente" ) = requestClean( "idgente" )
	session( "digest" ) = requestClean( "digest" )
%>
<html>

<head>
<title>Índice de apartados</title>
</head>

<frameset framespacing="2" border="0" cols="40%,*" frameborder="0">
  <frame name="izquierda" scrolling="auto" target="_self" src="http://www.istas.net/risctox/eco_tema_arbol_frames.asp" marginwidth="0" marginheight="0">
  <frame name="contenido" scrolling="auto" target="_self" src="http://www.istas.net/risctox/eco_editarpagina.asp" marginwidth="0" marginheight="0" style="border-left-style: dotted; border-width: 1; color:#000000">
</frameset>

</html>
