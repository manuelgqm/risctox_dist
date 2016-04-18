<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->

<%
response.End()
'sql = "select * from RQ_temporal_pesticidas"
sql = "select top 200 * from rq_temporal_pesticidas"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table>"
while not objr.eof
	
	
	'Compruebo si ese numero CAS está en la BBDD
	sql2 = "select count(id) as conta from dn_risc_sustancias where num_cas='"&trim(objr("numero_cas"))&"'"
	Set objr2 = Server.CreateObject ("ADODB.Recordset")
	objr2.Open sql2,objconn1,adOpenKeyset
	if objr2("conta")>0 then
		'response.write "<td>SI</td>"
	else
		response.write "<tr>"
		response.write "<td>"&objr("numero_cas")&"</td>"
		response.write "<td>NO</td>"
		response.write "<td>"&objr("nombre")&"</td>"
		response.write "</tr>"
	end if
	
	
			
	
	objr.movenext
wend
response.write "</table>"


%>
