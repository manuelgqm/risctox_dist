<!--#include file="../EliminaInyeccionSQL.asp"-->
<%

 	Const adOpenKeyset = 1
	DIM objConnection
	DIM objRecordset

	idgente = session("idgente")

	if cstr(idgente)="" then response.redirect "error.htm"

	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	'OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=disoltec02; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
	OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
	'strConn = "DRIVER={SQL Server};Description=conexionsql;SERVER=192.168.0.150,1433;UID=spl;PWD=***REMOVED***;DATABASE=istas_risctox;"
	OBJConnection.connectionstring=strConn
	OBJConnection.Open

%>