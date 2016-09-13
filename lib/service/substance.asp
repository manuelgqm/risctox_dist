<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID="1034"%>	
<!--#include file="../urlManipulations.asp"-->
<!--#include file="../JSON151.asp"-->
<!--#include file="../class/SubstanceClass.asp"-->
<!--#include file="../../config/dbConnection.asp"-->
<!--#include file="../../dn_funciones_texto.asp"-->
<!--#include file="../../dn_funciones_comunes.asp"-->

<%
dim substanceId, substanceFields, mySubstance
Server.ScriptTimeout = 600
response.expires = -1

substanceId = obtainSanitizedQueryParameter("substanceId")
id_sustancia = substanceId
set mySubstance = new SubstanceClass
mySubstance.find substanceId, objConnection2
set substanceFields = mySubstance.fields

response.write ((new JSON).toJSON("data", substanceFields, false))
response.end
%>