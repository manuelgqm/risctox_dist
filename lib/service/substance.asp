<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID="1034"%>
<!--#include file="../urlManipulations.asp"-->
<!--#include file="../inputSanitizers.asp"-->
<!--#include file="../JSON151.asp"-->
<!--#include file="../class/SubstanceClass.asp"-->
<!--#include file="../db/substancesSearch.asp"-->
<!--#include file="../../config/dbConnection.asp"-->
<!--#include file="../dn_funciones_texto_utf-8.asp"-->
<!--#include file="../dn_funciones_comunes_utf-8.asp"-->
<!--#include file="../dictionaryManipulations.asp"-->

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
	mySubstance.obtainLevelOneFields substanceId, objConnection2
	dim i, key
	dim dictKeys : dictKeys = mySubstance.fields.Keys
	for i = 0 to Ubound(dictKeys)
		key = dictKeys(i)
		if not(hasValue(mySubstance.fields.item(key))) then 
			mySubstance.fields.remove(key)
		end if
	next

	set find = mySubstance.fields
end function

function findSalud()
	dim result : set result = Server.CreateObject("Scripting.Dictionary")
	result.add "grupo_iarc", "2A"
	result.add "volumen_iarc", "(VOL. 23, SUPL.7; 1987)"
	result.add "notas_iarc", Array()
	set findSalud = result
end function

function search()
	dim result : set result = Server.CreateObject("Scripting.Dictionary")
	dim name : name = obtainSanitizedQueryParameter("name")
	dim code : code = obtainSanitizedQueryParameter("code")
	dim searchType : searchType = getSearchType(name)
	name = replace(name, "*", "")
	dim searchQuery : searchQuery = obtainSearchQuery(name, code, searchType)
	dim substancesRecordset : Set substancesRecordset = Server.CreateObject("ADODB.Recordset")
	const adOpenStatic = 3
	const adCmdText = 1
	substancesRecordset.Open searchQuery, objConnection2, adOpenStatic, adCmdText

	result.add "name", name
	result.add "records", substancesRecordset
	set search = result
end function

' PRIVATE '
function getSearchType(name)
	dim  result : result = "exacto"
	if hasArterisk(name) then result = ""
	getSearchType = result
end function

function hasArterisk(str)
	hasArterisk = false
	hasArterisk = inStr(str, "*") > 0
end function
%>