<%
' Detectamos el navegador
' Y establecemos tamaños de campos

user_agent = Request.ServerVariables("HTTP_USER_AGENT")

if instr(user_agent,"MSIE") then
	navegador = "ie"
	len_auto_prod_nombre="55"
	len_auto_comp_nombre="90"
elseif instr(user_agent,"Mozilla") then
	navegador = "mozilla"
	len_auto_prod_nombre="55"
	len_auto_comp_nombre="80"
elseif instr(user_agent,"Opera") then
	navegador = "opera"
	len_auto_prod_nombre="45"
	len_auto_comp_nombre="70"
else ' desconocido
	navegador = ""
	len_auto_prod_nombre="55"
	len_auto_comp_nombre="70"
end if
%>
