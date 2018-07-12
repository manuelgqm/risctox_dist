<%
Const adOpenKeyset = 1
DIM objConnection
DIM objRecordset

FUNCTION unquote(s)
  pos = instr(s, "'")
  while pos > 0 
    s = mid(s,1,pos) & "'" & mid(s,pos+1)
    pos = instr(pos+2, s, "'")
  wend
  pos = instr(s, """") 
  while pos > 0 
    s = mid(s,1,pos-1) & "''" & mid(s,pos+1)
    pos = instr(pos+2, s, """")
  wend
  unquote = trim(s)
end FUNCTION

Set OBJConnection = Server.CreateObject("ADODB.Connection")
'OBJConnection.Open "driver={sql server};server=disoltec02;database=istas_risctox;UID=xip_web;PWD=***REMOVED**"
OBJConnection.Open "driver={sql server};server=DISOLTEC03\XIP;database=istas_risctox;UID=xip_web;PWD=***REMOVED**"

url_antigua = "http://www.mtas.es/insht"
url_nueva = "http://www.insht.es"

sql = "SELECT id,definicion FROM RQ_DEFINICIONES WHERE definicion LIKE '%"&url_antigua&"%'"
set dSQL = Server.CreateObject ("ADODB.Recordset")
set dSQL = OBJConnection.Execute(sql)
do while not dSQL.eof
	id = dSQL("id")
	definicion = dSQL("definicion")
	definicion = unquote(definicion)
	definicion_nueva = replace(definicion,url_antigua,url_nueva)
	response.write id&","
	sql2 = "UPDATE RQ_DEFINICIONES SET definicion='"&definicion_nueva&"' WHERE id="&id
	set dSQL2 = Server.CreateObject ("ADODB.Recordset")
	set dSQL2 = OBJConnection.Execute(sql2)
	dSQL.movenext
loop



%>