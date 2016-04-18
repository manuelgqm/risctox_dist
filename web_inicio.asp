<!--#include file="EliminaInyeccionSQL.asp"-->
<%
Const adOpenKeyset = 1
DIM objConnection
DIM objRecordset

Set OBJConnection = Server.CreateObject("ADODB.Connection")
OBJConnection.Open "Provider=SQLOLEDB; Data Source=osiris.servidoresdns.net; Initial Catalog=qc507; User ID=qc507; Password=sql"

%>