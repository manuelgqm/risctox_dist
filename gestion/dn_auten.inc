<%
if session("usuario")="" then
	flashMsgCreate "Su sesi�n ha cadudado. Por favor introduzca de nuevo sus claves", "Advertencia"
	response.redirect ("dn_index.asp")
end if
%>
