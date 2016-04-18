<%
if session("usuario")="" then
	flashMsgCreate "Su sesión ha cadudado. Por favor introduzca de nuevo sus claves", "Advertencia"
	response.redirect ("dn_index.asp")
end if
%>
