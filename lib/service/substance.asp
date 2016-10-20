<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID="1034"%>
<!--#include file="../urlManipulations.asp"-->
<!--#include file="../inputSanitizers.asp"-->
<!--#include file="../JSON151.asp"-->
<!--#include file="../class/SubstanceClass.asp"-->
<!--#include file="../../config/dbConnection.asp"-->
<!--#include file="../../dn_funciones_texto.asp"-->
<!--#include file="../../dn_funciones_comunes.asp"-->

<%
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Server.ScriptTimeout = 600
response.expires = -1

dim action : action = obtainSanitizedQueryParameter("action")
dim JSONResponse : JSONResponse = execute( sanitizeScript(action) & "()" )

function find()
	dim substanceId, substanceFields, mySubstance
	substanceId = obtainSanitizedQueryParameter("substanceId")
	id_sustancia = substanceId
	set mySubstance = new SubstanceClass
	mySubstance.find substanceId, objConnection2
	set substanceFields = mySubstance.fields
	response.write ((new JSON).toJSON("data", substanceFields, false))
end function

function search()
	dim name : name = obtainSanitizedQueryParameter("name")
	response.write name
	search = name
end function
%>