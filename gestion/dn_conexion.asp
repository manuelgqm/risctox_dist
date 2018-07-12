<%
' PARA ECHAR A QUIEN ESTE DENTRO EN CASO DE NECESIDAD
'session.abandon

' ############################
' MODO: PRUEBAS O PRODUCCION
' ############################
'modo = "pruebas"
modo = "produccion"
session("modo")=modo

if (modo = "produccion") then
  'on error resume next
end if



dim objConn1
Set objConn1 = Server.CreateObject("ADODB.Connection")
'strConn = "DRIVER={SQL Server};Description=conexionsql;SERVER=cronos.spl-ssi.com,1433;UID=spl;PWD=***REMOVED**;DATABASE=istas_risctox;"
strConn = "Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED**"


if (modo = "pruebas") then
  objConn1.connectionstring="Provider=SQLOLEDB; Data Source=lwda680.servidoresdns.net; Initial Catalog=qbk243; User ID=qbk243; Password=***REMOVED**"
elseif (modo = "produccion") then
  'objConn1.connectionstring="Provider=SQLOLEDB; Data Source=osiris.servidoresdns.net; Initial Catalog=qc507; User ID=qc507; Password=sql"
  'objConn1.connectionstring="Provider=SQLOLEDB; Data Source=217.13.81.22; Initial Catalog=istas_formacion; User ID=xip_web; Password=XiPmm7337"
  'objConn1.connectionstring="Provider=SQLOLEDB; Data Source=217.13.81.22; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED**"
'  objConn1.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED**"
'  objConn1.connectionstring="Provider=SQLOLEDB; Data Source=NEPTUNO_HP\SQLEXPRESS; Initial Catalog=istas_risctox; User ID=spl; Password=***REMOVED**"
	objConn1.connectionstring=strConn

end if

' ############################################################
' CAMBIO DEL TIEMPO DE CONEXION POR DEFECTO, PARA ALARGARLO
segundos=72000
Server.ScriptTimeOut=segundos 'tiempo de script
'tiempo de conexion: debemos establecerlo antes de abrir
objConn1.ConnectionTimeout=segundos 
objConn1.CommandTimeout=segundos

' ES NECESARIO PARA QUE DOS LECTURAS CONSECUTIVAS DE UN RECORDSET NO VACÍE EL MISMO
objConn1.CursorLocation = adUseClient

' ############################################################

objConn1.Open
%>

<%
sub cerrarconexion
	objConn1.close
	Set objConn1=nothing
end sub
%>
