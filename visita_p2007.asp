<% 
   'idgente = session("idgente")
   'if cstr(idgente)="" then response.redirect "panelcontrol.asp?texto=Se ha pasado el tiempo de espera. Vuelve a identificarte"
   session("id_ecogente")=""
%>

<html>

<head>
<title>Visita virtual RISCTOX</title>
</head>

<frameset framespacing="0" border="0" cols="75,*" frameborder="0">
  <frame name="lateral" scrolling="no" src="visitavirtual_vertical2007.html">
  <frame name="cuerpo_visita" src="visita_paso2007.asp?paso=1" target="_self">
</frameset>

</html>
