<%
function obtainDefinitions(byVal keysSrz, connection)
	optainDefinitions = Array()
	Dim keys
	if isNull(keysSrz) or keysSrz = "" then 
		exit function
	end if
	keys = split(keysSrz, ",")
	dim i, key
	dim keysFormated : keysFormated = array()
	for i = 0 to Ubound(keys)
		key = replace(keys(i), " ", "")
		if len(key) then arrayPush keysFormated, key
	next
	obtainDefinitions = findDefinitions(keysFormated, connection)
end function

function findDefinitions(keys, connection)
	Dim result : result = Array()
	dim keysQueryFormatedSrz : keysQueryFormatedSrz =  formatKeysQueryList(keys)
	Dim definitionsQuery : definitionsQuery = composeQuery(keysQueryFormatedSrz)
	if keysQueryFormatedSrz = "" then
		findDefinitions = result
		exit function
	end if
	Dim definitionsRecordset : set definitionsRecordset = connection.execute(definitionsQuery)
	Dim i, definitionId, definitionText, key
	For i = 0 to Ubound(keys)
		set key = Server.CreateObject("Scripting.Dictionary")
		definitionId = ""
		definitionText = ""
		if not definitionsRecordset.EOF then
			definitionId = definitionsRecordset("id").value
			definitionText = definitionsRecordset("definicion").value
			definitionsRecordset.moveNext
		end if
		key.add "id", definitionId
		key.add "description", definitionText
		key.add "key", keys(i)
		result = arrayPushDictionary(result, key)
	next

	definitionsRecordset.close
	set definitionsRecordset = nothing
	findDefinitions = result
end function

function formatKeysQueryList(byVal keys)
	dim result : result = ""
	dim format1 : format1 = Array("1", "2", "3", "4", "5", "6", "7", "8", "o")
	dim format2 : format2 = Array("F", "I", "S")
	dim i, definitionKeyFormated, key
	for i = 0 to ubound(keys)
		key = trim(keys(i))
		definitionKeyFormated = key
		if inArray(key, format1) then
			definitionKeyFormated = "(" & key & ")"
		end if
		if inArray(key, format2) then
			definitionKeyFormated = lcase(key) & "."
		end if
		keys(i) = definitionKeyFormated
	next
	keys = arrayWrapItems(keys, "'", "'")
	result = join(keys, ",")
	formatKeysQueryList = result
end function

function composeQuery(keysQueryFormatedSrz)
	Dim sql
	sql = "SELECT id, palabra, dbo.udf_StripHTML(definicion) as definicion FROM rq_definiciones where palabra in (" & keysQueryFormatedSrz & ")"
	composeQuery = sql
end function
%>