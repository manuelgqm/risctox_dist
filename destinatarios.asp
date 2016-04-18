<%
 	Const adOpenKeyset = 1
	DIM objConnection	
	DIM objRecordset
	
	Set OBJConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
	
	sql = "SELECT destinatarios FROM BEL_ENVIADOS WHERE idboletin=165"
	Set objRecordset = Server.CreateObject ("ADODB.Recordset")
	set objRecordset = OBJConnection.Execute(sql)
	destinatarios = objRecordset("destinatarios")
	response.write replace(destinatarios,";","<br>")
	
%>