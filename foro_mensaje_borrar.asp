<!--#include file="EliminaInyeccionSQL.asp"-->
<%

iden = session("id_ecogente")
valor = EliminaInyeccionSQL(request("id"))

Const adOpenKeyset = 1
DIM objConnection
DIM objRecordset

Set OBJConnection = Server.CreateObject("ADODB.Connection")
OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

SQLQuery = "SELECT * FROM ECO_FOROS WHERE id="&valor
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
set objRecordset = OBJConnection.Execute(SQLQuery)

sigue = objRecordSet("sig")
sqlquery4 = "UPDATE ECO_FOROS SET sig="&sigue&" WHERE sig="&valor
set objRecordset4 = OBJConnection.Execute(SQLQuery4)

sqlquery5 = "UPDATE ECO_FOROS SET tipo=0 WHERE id="&valor
set objRecordset5 = OBJConnection.Execute(SQLQuery5)


%>

<html>

<head><title>Borrar anuncio</title></head>


<body bgcolor="#FFFFFF" onLoad="javascript:vamos()">
ok, mensaje borrado
</body>
</html>

<script LANGUAGE="JScript">

function vamos() {
window.opener.location.reload()
window.close();

}

</script>
