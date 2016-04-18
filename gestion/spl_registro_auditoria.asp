<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="../dn_funciones_texto.asp" 'necesario para el corte de texto -->
<!--#include file="dn_auten.inc"-->

<%
'si nos pasan id de registro, consultamos datos
id=EliminaInyeccionSQL(request("id"))
%>
	<!--#include file="adovbs.inc"-->
	<!--#include file="dn_conexion.asp"-->
<%
if id<>"" then

	Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
	PID = "PID=" & UploadProgress.CreateProgressID()
	barref = "arbol_framebar.asp?to=10&" & PID

	'DATOS GENERALES
	
	sql3="select * from spl_auditoria where id=" &id
	'response.write sqls
	set objRst3=objconn1.execute(sql3)
	
	'id=objRst3("id")
	usuario=objRst3("usuario")
	fecha=objRst3("fecha")
	ip=objRst3("ip")
	navegador=objRst3("navegador")
	entidad=objRst3("entidad")
	accion=objRst3("accion")
	descripcion=objRst3("descripcion")

	objRst3.close
	set objRst3=nothing
	

end if
%>

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
Nifty("div#box2","big"); 
}

</script>
</head>

<body>
<%flashMsgShow()%>
<div id="box2" class="centcontenido">
<h2>Ficha de registro de auditor&iacute;a</h2>
	<table>
		<tr>
			<td><b>Fecha</b></td>
			<td>&nbsp;</td>
			<td><%=fecha%></td>
		</tr>
		<tr>
			<td><b>Usuario</b></td>
			<td>&nbsp;</td>
			<td><%=usuario%></td>
		</tr>
		<tr>
			<td><b>IP</b></td>
			<td>&nbsp;</td>
			<td><%=ip%></td>
		</tr>
		<tr>
			<td><b>Navegador</b></td>
			<td>&nbsp;</td>
			<td><%=navegador%></td>
		</tr>
		<tr>
			<td><b>Acci&oacute;n</b></td>
			<td>&nbsp;</td>
			<td><%=accion%></td>
		</tr>
		<tr>
			<td><b>Entidad</b></td>
			<td>&nbsp;</td>
			<td><%=entidad%></td>
		</tr>
		<tr>
			<td><b>Descripci&oacute;n complementaria</b></td>
			<td>&nbsp;</td>
			<td><%=descripcion%></td>
		</tr>
	</table>
	
</div>
</body>
</html>
<%
	cerrarconexion
%>