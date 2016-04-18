<%
' CONEXIÓN A LAS BASES DE DATOS.
' Hay dos conexiones, una a la base antigua y otra a la nueva.

dim objConnection, objConnection2
	
' ############################################################
' ### Base antigua
' ############################################################
set objConnection = Server.CreateObject("ADODB.Connection")
'objConnection.open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("../sql.mdb")
'OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
 OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_formacion; User ID=istas_sql_usuari; Password=***REMOVED***"

' ############################################################
' ### Base nueva
' ############################################################
set objConnection2 = Server.CreateObject("ADODB.Connection")
objConnection2.connectionstring="Provider=SQLOLEDB; Data Source=lwda680.servidoresdns.net; Initial Catalog=qbk243; User ID=qbk243; Password=***REMOVED***"
' CAMBIO DEL TIEMPO DE CONEXION POR DEFECTO, PARA ALARGARLO
segundos=120
Server.ScriptTimeOut=segundos 'tiempo de script
'tiempo de conexion: debemos establecerlo antes de abrir
objConnection2.ConnectionTimeout=segundos 
objConnection2.CommandTimeout=segundos

objConnection2.Open
%>

<%
sub cerrarconexion
	' Cierra las dos conexiones
	objConnection.close
	set objConnection=nothing
	objConnection2.close
	set objConnection2=nothing
end sub
%>
