<%
' ############################
' MODO: PRUEBAS O PRODUCCION
' ############################
'modo = "pruebas"
modo = "produccion"
session("modo")=modo

if (modo = "produccion" or modo = "pruebas") then
  ' on error resume next
end if

' BASES DE DATOS
dim objConnection, objConnection2
	
' ############################################################
' ### Base antigua (siempre la misma)
' ############################################################
set objConnection = Server.CreateObject("ADODB.Connection")
'objConnection.open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("../sql.mdb")
'OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
'OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=osiris.servidoresdns.net; Initial Catalog=qc507; User ID=qc507; Password=sql"
'OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=disoltec02; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
' OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
' OBJConnection.connectionstring="driver={sql server};server=HP-LOLO\SQLEXPRESS;database=istas_risctox;UID=istas_SQL;PWD=***REMOVED***"

Set wshShell = CreateObject( "WScript.Shell" )
localConnectionString = wshShell.ExpandEnvironmentStrings( "%istas_risctox_dbConnectionString%" )
set wshShell = Nothing

if localConnectionString <> "" then
	connectionString = localConnectionString
else
	connectionString = "Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
end if

OBJConnection.connectionstring = connectionString
OBJConnection.Open

set objConnection2 = Server.CreateObject("ADODB.Connection")
OBJConnection2.connectionstring = connectionString

' CAMBIO DEL TIEMPO DE CONEXION POR DEFECTO, PARA ALARGARLO
segundos=120
Server.ScriptTimeOut=segundos 'tiempo de script
'tiempo de conexion: debemos establecerlo antes de abrir
objConnection2.ConnectionTimeout=segundos 
objConnection2.CommandTimeout=segundos


' ES NECESARIO PARA QUE DOS LECTURAS CONSECUTIVAS DE UN RECORDSET NO VACÍE EL MISMO
' objConnection2.CursorLocation = adUseClient

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