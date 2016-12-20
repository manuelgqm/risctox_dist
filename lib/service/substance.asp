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

function findSection()
	dim section : section = obtainSanitizedQueryParameter("section")
	dim substanceId : substanceId = obtainSanitizedQueryParameter("substanceId")
	select case(section)
		case("identificacion"):
			set findSection = obtainIdentificacionFields(substanceId)
		case("salud"):
			set findSection = obtainSaludFields(substanceId)
		case("medioAmbiente"):
			set findSection = obtainMedioAmbienteFields(substanceId)
		case else:
			set findSection = Server.CreateObject("Scripting.Dictionary")
	end select
end function

function findCancerOtras()
	dim substanceId : substanceId = obtainSanitizedQueryParameter("substanceId")
	dim substance : set substance = new SubstanceClass
	substance.obtainCancerOtrasFields substanceId, objConnection2
	
	set findCancerOtras = removeDictionaryEmptyFields(substance.fields)	
end function

function findEnfermedades()
	dim substanceId : substanceId = obtainSanitizedQueryParameter("substanceId")
	dim substance : set substance = new SubstanceClass
	substance.obtainEnfermedadesFields substanceId, objConnection2
	set findEnfermedades = substance.fields
end function

function obtainIdentificacionFields(substanceId)
	dim substance : set substance = new SubstanceClass
	substance.obtainLevelOneFields substanceId, objConnection2
	
	set obtainIdentificacionFields = removeDictionaryEmptyFields(substance.fields)
end function

function obtainSaludFields(substanceId)
	dim substance : set substance = new SubstanceClass
	substance.obtainSaludFields substanceId, objConnection2
	set obtainSaludFields = removeDictionaryEmptyFields(substance.fields)
end function

function obtainMedioAmbienteFields(substanceId)
	dim substance : set substance = new SubstanceClass
	substance.obtainMedioAmbienteFields substanceId, objConnection2
	set obtainMedioAmbienteFields = removeDictionaryEmptyFields(substance.fields)
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

function removeDictionaryEmptyFields(byRef dictionary)
	dim result : set result = Server.CreateObject("Scripting.Dictionary")
	dim i, key
	dim dictKeys : dictKeys = dictionary.Keys
	for i = 0 to Ubound(dictKeys)
		key = dictKeys(i)
		if not(hasValue(dictionary.item(key))) then 
			dictionary.remove(key)
		end if
	next

	set removeDictionaryEmptyFields = dictionary
end function
%>