<% 
   'idgente = session("idgente")
   'if cstr(idgente)="" then response.redirect "panelcontrol.asp?texto=Se ha pasado el tiempo de espera. Vuelve a identificarte"
%>

<html>

<head>
<title>Visita virtual ECOinformas</title>
</head>

<frameset framespacing="0" border="0" cols="75,*" frameborder="0">
  <frame name="lateral" scrolling="no" src="visitavirtual_vertical_p.html">
  <frame name="cuerpo_visita" src="visita_paso_p.asp?paso=1" target="_self">
</frameset>

</html>
