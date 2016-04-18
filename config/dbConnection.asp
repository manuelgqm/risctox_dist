<!--#include file="dbConnectionStrings.asp"-->
<%
' CONEXIÓN A LAS BASES DE DATOS.
' Hay dos conexiones, una a la base antigua y otra a la nueva.
' SERVER=lwda329.servidoresdns.net

'modo = "pruebas"
modo = "produccion"
session("modo")=modo

if (modo = "produccion") then
  'on error resume next
end if

' BASES DE DATOS
dim objConnection, objConnection2
	
set objConnection = Server.CreateObject("ADODB.Connection")
OBJConnection.connectionstring = DB_CONNECTION_STRING_DEV

set objConnection2 = Server.CreateObject("ADODB.Connection")

if (modo = "pruebas") then
  objConnection2.connectionstring="Provider=SQLOLEDB; Data Source=lwda680.servidoresdns.net; Initial Catalog=qbk243; User ID=qbk243; Password=***REMOVED***"
elseif (modo = "produccion") then
  OBJConnection2.connectionstring = DB_CONNECTION_STRING_DEV
end if

connStr = DB_CONNECTION_STRING_DEV

Server.ScriptTimeOut = 120 'seconds
objConnection2.ConnectionTimeout = segundos 
objConnection2.CommandTimeout = segundos
objConnection2.Open
%>

<%
sub cerrarconexion
	objConnection.close
	set objConnection = nothing
	objConnection2.close
	set objConnection2 = nothing
end sub
%>
