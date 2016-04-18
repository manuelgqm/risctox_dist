<%

Const adOpenKeyset = 1
DIM objConnection	
DIM objRecordset

Set OBJConnection = Server.CreateObject("ADODB.Connection")
'OBJConnection.Open "driver={sql server};server=osiris.servidoresdns.net;database=qc507;UID=qc507;PWD=sql"
OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"


Session.TimeOut = 1000
clave = trim(request("clave"))
contra = trim(request("contra"))

orden = "SELECT idgente FROM ECOINFORMAS_GENTE WHERE clave='"&clave&"' AND (contra='"&contra&"' OR email='"&contra&"')"
Set dSQL = Server.CreateObject ("ADODB.Recordset")
dSQL.Open orden,objConnection,adOpenKeyset

if not(dSQL.bof and dSQL.eof) and clave<>"" and contra<>"" then

	'-- Crear variables sesión y pasarlas a la página de ENTRADA
	session("id_ecogente") = dSQL("idgente")

	'-- ABrir la página
	pagina_esp = request("pagina_esp")
	if pagina_esp="" then
		if request("idenlace")="" then
			response.redirect ("index.asp?idpagina="&request("idpagina"))
		else
			response.redirect ("abreenlacer.asp?idenlace="&request("idenlace"))
		end if
	else
		response.redirect (pagina_esp)
	end if
		
	
else
	
	session("id_ecogente") = ""
	response.redirect ("acceso.asp?error=1")
	
end if
%>