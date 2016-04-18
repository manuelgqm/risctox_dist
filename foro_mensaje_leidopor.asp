<!--#include file="EliminaInyeccionSQL.asp"-->
<%Response.Expires=0%>
<%
iden = session("id_ecogente")
valor = EliminaInyeccionSQL(request("id"))

Const adOpenKeyset = 1
DIM objConnection
DIM objRecordset

Set OBJConnection = Server.CreateObject("ADODB.Connection")
OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"
 	
sqlquery = "SELECT ECO_FOROS_LEIDOS.fecha,ECOINFORMAS_GENTE.nombre,ECOINFORMAS_GENTE.apellidos FROM ECO_FOROS_LEIDOS LEFT JOIN ECOINFORMAS_GENTE ON ECO_FOROS_LEIDOS.idgente=ECOINFORMAS_GENTE.idgente WHERE ECO_FOROS_LEIDOS.idmensaje="&valor&" ORDER BY ECO_FOROS_LEIDOS.idgente,ECO_FOROS_LEIDOS.fecha DESC"
Set objRecordset = Server.CreateObject ("ADODB.Recordset")
objrecordset.Open SQLQuery,objConnection,adOpenKeyset
numvisitas = objrecordset.recordcount
%>
<html>

<head>
<title>ACCESOS a este mensaje</title>
<link rel="stylesheet" type="text/css" href="estructura.css">
</head>

<body bgcolor="#FFFFFF" topmargin="10" leftmargin="10" class="cuerpo">
<p class="texto">
Estos son los accesos a este mensaje:
</p>
<table border="0" cellpadding="2" cellspacing="2" width="100%" class="tabla">
   <tr>
     <td class="celda2" align="center">Nº</td>
     <td class="celda2" align="center">PERSONA</td>
     <td class="celda2" align="center">FECHA</td>
   </tr>
<% 
v = -1
fecha = ""
do while not objrecordset.eof 
  v = v+1
  if fecha<>objrecordset("fecha") then
  	estilo = "font-weight: bold"
  else
  	estilo = "font-weight: normal"
  end if
%>
   <tr>
     <td class="celda2" align="left"><%=numvisitas-v%></td>
     <td class="celda2" align="left"><%=objrecordset("nombre")&" "&objrecordset("apellidos")%></td>
     <td class="celda2" style="<%=estilo%>" align="left"><%=objrecordset("fecha")%></td>
   </tr>
<%
  fecha = objrecordset("fecha")
  objrecordset.movenext
loop
%>
</table>

</body>
</html>
