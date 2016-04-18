<!--#include file="../../EliminaInyeccionSQL.asp"-->
<!--#include file="../dn_conexion.asp"-->

<%
Session.TimeOut = 1000
clave = trim(request("clave"))
clave = EliminaInyeccionSQL(clave)
contra = trim(request("contra"))
contra = EliminaInyeccionSQL(contra)
idpagina_request = request("idpagina")
idpagina_request = EliminaInyeccionSQL(idpagina_request)
idenlace_request = request("idenlace")
idenlace_request = EliminaInyeccionSQL(idenlace_request)

orden = "SELECT idgente FROM ECOINFORMAS_GENTE_NUEVO WHERE clave='"&clave&"' AND (contra='"&contra&"' OR email='"&contra&"')"
Set dSQL = Server.CreateObject ("ADODB.Recordset")
dSQL.Open orden,objConnection,adOpenKeyset

if not(dSQL.bof and dSQL.eof) and clave<>"" and contra<>"" then

	'-- Crear variables sesión y pasarlas a la página de ENTRADA
	session("id_ecogente2") = dSQL("idgente")
	'Sergio
	idpagina_request = 963
	'-- ABrir la página
	pagina_esp = request("pagina_esp")
	pagina_esp = EliminaInyeccionSQL(pagina_esp)
	if pagina_esp="" then
		if idenlace_request="" then
			response.redirect ("index.asp?idpagina="&idpagina_request)
		else
			response.redirect ("abreenlacer.asp?idenlace="&idenlace_request)
		end if
	else
		response.redirect (pagina_esp)
	end if
		
	
else
	
	session("id_ecogente2") = ""
	response.redirect ("acceso.asp?error=1")
	
end if
%>