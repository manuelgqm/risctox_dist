<!--#include file="web_inicio.asp"-->

<% 
	'Const adOpenKeyset = 1
	'DIM objConnection	
	'DIM objRecordset

	idgente = session("idgente")
	if cstr(idgente)="" Then idgente = "0"

%>
<html>

<head>
<title>Vigilancia Legislativa</title>
</head>

<frameset framespacing="2" border="0" cols="37%,*" frameborder="0">
  <frame name="izquierda" scrolling="auto" target="_self" src="vig_leg_listado.asp" marginwidth="0" marginheight="0">
  <frame name="derecha" scrolling="auto" target="_self" src="vig_leg_editar.asp" marginwidth="0" marginheight="0" style="border-left-style: dotted; border-width: 1; color:#000000">
</frameset>

</html>
