<!--#include file="EliminaInyeccionSQL.asp"-->
<%
' CONEXIÓN A LAS BASES DE DATOS.
' Hay dos conexiones, una a la base antigua y otra a la nueva.
' SERVER=lwda329.servidoresdns.net

' ############################
' MODO: PRUEBAS O PRODUCCION
' ############################
'modo = "pruebas"
modo = "produccion"
session("modo")=modo

if (modo = "produccion") then
  'on error resume next
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
OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"


' ############################################################
' ### Base nueva (depende de pruebas o produccion)
' ############################################################
set objConnection2 = Server.CreateObject("ADODB.Connection")

if (modo = "pruebas") then
  objConnection2.connectionstring="Provider=SQLOLEDB; Data Source=lwda680.servidoresdns.net; Initial Catalog=qbk243; User ID=qbk243; Password=***REMOVED***"
elseif (modo = "produccion") then
  'usuario: qc507
  'contraseña: sql
  'servidor: osiris.servidoresdns.net

  'objConnection2.connectionstring="Provider=SQLOLEDB; Data Source=osiris.servidoresdns.net; Initial Catalog=qc507; User ID=qc507; Password=sql"
  'OBJConnection2.connectionstring="Provider=SQLOLEDB; Data Source=disoltec02; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
  OBJConnection2.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED***"
end if


' CAMBIO DEL TIEMPO DE CONEXION POR DEFECTO, PARA ALARGARLO
segundos=120
Server.ScriptTimeOut=segundos 'tiempo de script
'tiempo de conexion: debemos establecerlo antes de abrir
objConnection2.ConnectionTimeout=segundos 
objConnection2.CommandTimeout=segundos


' ES NECESARIO PARA QUE DOS LECTURAS CONSECUTIVAS DE UN RECORDSET NO VACÍE EL MISMO
'objConnection2.CursorLocation = adUseClient

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
