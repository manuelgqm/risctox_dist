<!--#include file="../EliminaInyeccionSQL.asp"-->
<!--#include file="../_framework/common/MD5.asp"-->
<%

 	Const adOpenKeyset = 1
	DIM objConnection
	DIM objRecordset
	Dim key, sDigest, digest
	
	key = "szhWmCDkP5zp4"
	' sDigest = MD5( session("idgente") & key & left( cstr( time() ), 5 ) )
	sDigest = MD5( session("idgente") & key & request.serverVariables( "REMOTE_ADDR" ) )
	
	if session( "digest" ) <> sDigest or session( "idgente" ) = "" then 
		response.write( "Acceso no permitido" )
		response.end()
	end if

	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
	OBJConnection.connectionstring=strConn
	OBJConnection.Open

%>