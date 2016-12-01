<%
function findDefinitions(notas, connection)
	Dim result : result = Array()
	dim definitionKeysQueryFormatedSrz : definitionKeysQueryFormatedSrz =  formatDefinitionKeysQueryList(notas)
	Dim notasDefinitionsQuery : notasDefinitionsQuery = composeDefinitionsQuery(definitionKeysQueryFormatedSrz)
	if definitionKeysQueryFormatedSrz = "" then
		findDefinitions = result
		exit function
	end if
	Dim definitionsRecordset : set definitionsRecordset = connection.execute(notasDefinitionsQuery)
	Dim i, definitionId, definitionText, nota
	For i = 0 to Ubound(notas)
		set nota = Server.CreateObject("Scripting.Dictionary")
		definitionId = ""
		definitionText = ""
		if not definitionsRecordset.EOF then
			definitionId = definitionsRecordset("id").value
			definitionText = definitionsRecordset("definicion").value
			definitionsRecordset.moveNext
		end if
		nota.add "id", definitionId
		nota.add "description", definitionText
		nota.add "key", notas(i)
		result = arrayPushDictionary(result, nota)
	next

	definitionsRecordset.close
	set definitionsRecordset = nothing
	findDefinitions = result
end function

function formatDefinitionKeysQueryList(byVal definitionKeys)
	dim result : result = ""
	dim format1 : format1 = Array("1", "2", "3", "4", "5", "6", "7", "8", "o")
	dim format2 : format2 = Array("F", "I", "S")
	dim i, definitionKeyFormated, definitionKey
	for i = 0 to ubound(definitionKeys)
		definitionKey = trim(definitionKeys(i))
		definitionKeyFormated = definitionKey
		if inArray(definitionKey, format1) then
			definitionKeyFormated = "(" & definitionKey & ")"
		end if
		if inArray(definitionKey, format2) then
			definitionKeyFormated = lcase(definitionKey) & "."
		end if
		definitionKeys(i) = definitionKeyFormated
	next
	definitionKeys = arrayWrapItems(definitionKeys, "'", "'")
	result = join(definitionKeys, ",")
	formatDefinitionKeysQueryList = result
end function

function composeDefinitionsQuery(definitionKeysQueryFormatedSrz)
	Dim sql
	sql = "SELECT id, palabra, dbo.udf_StripHTML(definicion) as definicion FROM rq_definiciones where palabra in (" & definitionKeysQueryFormatedSrz & ")"
	composeDefinitionsQuery = sql
end function
%>