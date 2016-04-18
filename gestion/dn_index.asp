<!--#include file="dn_fun_comunes.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Istas</title>
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
window.onload=function(){
Nifty("div#box","big"); 
}
</script>
</head>

<body>
<%flashMsgShow()%>
<div id="box" style="margin-top:100px;" class="centcontenido">
<form action="dn_password.asp" method="post">
<p>Usuario: <input name="usuario" type="text" size="10" maxlength="7" />
</p>
<p>Clave: <input name="clave" type="password" size="10" />
</p>
<p><input type="submit" value="Enviar" /></p>
</form>
</div>
</body>
</html>
