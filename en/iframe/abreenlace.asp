<!--#include file="../../EliminaInyeccionSQL.asp"-->

<%

Set OBJConnection = Server.CreateObject("ADODB.Connection")
OBJConnection.Open "driver={sql server};server=DISOLTEC03\XIP;database=istas_web;UID=xip_web;PWD=***REMOVED**"

idenlace = request("idenlace")
idenlace = EliminaInyeccionSQL(idenlace)


sql = "SELECT visitas,url,afiliacion FROM ENL_ENLACES WHERE id="&idenlace
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
set objRecordset = OBJConnection.Execute(sql)

visitas = objRecordset("visitas")
url = objRecordset("url")
afiliacion = objRecordset("afiliacion")

sql = "UPDATE ENL_ENLACES SET visitas="&visitas+1&" WHERE id="&idenlace
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
set objRecordset = OBJConnection.Execute(sql)

if cstr(session("id_ecogente"))<>"" then
	id_gente = session("id_ecogente")
else
	id_gente = 0
end if
IP = Request.ServerVariables("REMOTE_ADDR")
Set MiBrowser = Server.CreateObject("MSWC.BrowserType")
navegador = MiBrowser.Browser
orden = "INSERT INTO ENL_VISITAS (fecha,hora,IP,navegador,idenlace,idgente) VALUES ('"&date()&"','"&time()&"','"&IP&"','"&navegador&"','"&idenlace&"','"&id_gente&"')"
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
Set objRecordset = OBJConnection.Execute(orden)

response.redirect url
%>