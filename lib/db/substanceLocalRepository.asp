<!--#include file="../vbsJSON.asp"-->
<%
function findLocalJSONSubstance(file, substanceId)
	Dim result : set result = Server.CreateObject("Scripting.Dictionary")
	Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject") 
	Dim json : Set json = New VbsJson
	Dim str : str = fso.OpenTextFile(file).ReadAll
	set result = json.Decode(str)

	'TODO: filter substance in dictionary'
	set findLocalJSONSubstance = result("data")(7)
end function
%>