<%
function obtainValoresLimiteBiologico(substance, connection)
	Dim result : result = Array()
	Dim indicadores : indicadores = Array("ib_1", "ib_2", "ib_3", "ib_4", "ib_5", "ib_6")
	Dim valores : valores = Array("vlb_1", "vlb_2", "vlb_3", "vlb_4","vlb_5", "vlb_6")
	Dim momentos : momentos = Array("momento_1", "momento_2", "momento_3", "momento_4", "momento_5", "momento_6")
	Dim notas : notas = Array("notas_vlb_1", "notas_vlb_2", "notas_vlb_3", "notas_vlb_4","notas_vlb_5", "notas_vlb_6")
	Dim i, valorLimiteBiologico
	for i = 0 to ubound(indicadores)
		set valorLimiteBiologico = extractValorLimiteBiologico _
			( substance _
			, connection _
			, indicadores(i) _
			, valores(i) _
			, momentos(i) _
			, notas(i) _
			)
		if not isDictionaryEmpty(valorLimiteBiologico) then
			result = arrayPushDictionary(result, valorLimiteBiologico)
		end if
	next

	obtainValoresLimiteBiologico = result
end function

function extractValorLimiteBiologico(substance, connection, indicador, valor, momento, notas)
	Dim result : set result = Server.CreateObject("Scripting.Dictionary")
	result.add "indicador", substance.item(indicador)
	result.add "valor", substance.item(valor)
	result.add "momento", substance.item(momento)
	result.add "notas", obtainNotasValoresLimiteBiologico(substance.item(notas), connection)

	set extractValorLimiteBiologico = result
end function

function isDictionaryEmpty(dictionary)
	dim result : result = true
	if dictionary.count = 0 then
		isDictionaryEmpty = result
		exit function
	end if
	dim dictItems : dictItems = dictionary.items
	dim i, dictItem
	for i = 0 to ubound(dictItems)
		dictItem = dictItems(i)
		if isArray(dictItem) then
			if Ubound(dictItem) > -1 then
				isDictionaryEmpty = false
			end if
		else
			if dictItems(i) <> "" then
				isDictionaryEmpty = false
				exit function
			end if
		end if
	next

	isDictionaryEmpty = result
end function

function obtainNotasValoresLimiteBiologico(byVal notasSrz, connection)
	dim result : result = Array()
	Dim notas
	if isNull(notasSrz) then 
		notasSrz = ""
	end if
	notas = split(notasSrz, ",")
	dim definitionKeysQueryFormatedSrz : definitionKeysQueryFormatedSrz = formatDefinitionKeysQueryList(notas)
	obtainNotasValoresLimiteBiologico = findDefinitions(notas, definitionKeysQueryFormatedSrz, connection)
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

function findDefinitions(notas, definitionKeysQueryFormatedSrz, connection)
	Dim result : result = Array()
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

function composeDefinitionsQuery(definitionKeysQueryFormatedSrz)
	Dim sql
	sql = "SELECT id, palabra, dbo.udf_StripHTML(definicion) as definicion FROM rq_definiciones where palabra in (" & definitionKeysQueryFormatedSrz & ")"
	composeDefinitionsQuery = sql
end function
%>