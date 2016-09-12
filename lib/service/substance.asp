<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID="1034"%>	
<!--#include file="../urlManipulations.asp"-->
<!--#include file="../class/SubstanceClass.asp"-->
<!--#include file="../../config/dbConnection.asp"-->
<!--#include file="../../dn_funciones_texto.asp"-->
<!--#include file="../../dn_funciones_comunes.asp"-->

<%
dim substanceId, substance, mySubstance
Server.ScriptTimeout = 600
response.expires = -1

substanceId = obtainSanitizedQueryParameter("substanceId")
id_sustancia = substanceId
set mySubstance = new SubstanceClass
mySubstance.find substanceId, objConnection2
set substance = mySubstance.fields

response.write("{""name"": """ & substance.item("nombre") & """}")

'#PRIVATE'
Function FieldExists(ByVal rs, ByVal fieldName)

    On Error Resume Next
    FieldExists = rs.Fields(fieldName).name <> ""
    If Err <> 0 Then FieldExists = False
    Err.Clear

End Function
%>