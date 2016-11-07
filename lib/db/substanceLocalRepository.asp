<!--#include file="../vbsJSON.asp"-->
<%
function findLocalJSONSubstance(file, substanceId)
	Dim stream : set stream = Server.CreateObject("Scripting.Dictionary")
	Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject") 
	Dim json : Set json = New VbsJson
	Dim str : str = fso.OpenTextFile(file).ReadAll
	set stream = json.Decode(str)
	set findLocalJSONSubstance = findSubstanceInArray(stream("data"), substanceId)
end function

function findSubstanceInArray(arr, substanceId)
	set findSubstanceInArray = Server.CreateObject("Scripting.Dictionary")
	if not isArray(arr) then exit function
	dim i
	for i = 0 to uBound(arr)
		if arr(i).item("substanceId") = substanceId then
			set findSubstanceInArray = arr(i)
			exit function
		end if 
	next
end function
%>