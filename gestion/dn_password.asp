<!--#include file="dn_fun_comunes.asp"-->
<%
if request("intranet_istas")="si" or (request.form("usuario")="spl" and request.form("clave")="***REMOVED**") or ( request.form( "usuario" ) = "admin" and request.form( "clave" ) = "w9gzdm" ) then
	session.timeout = 480 ' Incrementamos tiempo de sesión a 8 horas
	session("usuario")="SPL"
	response.Redirect("dn_portada.asp")
else
	flashMsgCreate "Usuario y/o contrase&ntilde;a incorrectos", "Error"
	response.Redirect("dn_index.asp")
end if
%>
