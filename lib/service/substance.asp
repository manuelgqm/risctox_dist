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
dim actionResult
execute( _
	"set actionResult = " & sanitizeScript(action) & "()" _
)
response.write ((new JSON).toJSON("data", actionResult, false))

function find()
	dim substanceId, mySubstance
	substanceId = obtainSanitizedQueryParameter("substanceId")
	id_sustancia = substanceId
	set mySubstance = new SubstanceClass
	mySubstance.find substanceId, objConnection2
	
	set find = mySubstance.fields
end function

function search()
	dim result : set result = Server.CreateObject("Scripting.Dictionary")
	dim name : name = obtainSanitizedQueryParameter("name")
	result.add "name", name
	
	set search = result
end function
%>