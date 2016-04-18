<%
' INCLUIR ESTO EN LAS PAGINAS RESTRINGIDAS PARA EXPULSAR SI NO SE HA IDENTIFICADO
if session("id_ecogente")="" then
	'response.redirect "acceso.asp?idpagina=576"
	session("id_ecogente") = 12471
	session("risctox_en_webistas")="si"
end if
%>
