<%
function obtainValoresLimiteAmbiental(substance, connection)
	Dim result : result = Array()
	Dim estados : estados = Array("estado_1", "estado_2", "estado_3", "estado_4", "estado_5", "estado_6")
	Dim ed_ppm : ed_ppm = Array("vla_ed_ppm_1", "vla_ed_ppm_2", "vla_ed_ppm_3", "vla_ed_ppm_4","vla_ed_ppm_5", "vla_ed_ppm_6")
	Dim ed_mg_m3 : ed_mg_m3 = Array("vla_ed_mg_m3_1", "vla_ed_mg_m3_2", "vla_ed_mg_m3_3", "vla_ed_mg_m3_4", "vla_ed_mg_m3_5", "vla_ed_mg_m3_6")
	Dim ec_ppm : ec_ppm = Array("vla_ec_ppm_1", "vla_ec_ppm_2", "vla_ec_ppm_3", "vla_ec_ppm_4","vla_ec_ppm_5", "vla_ec_ppm_6")
	Dim ec_mg_m3 : ec_mg_m3 = Array("vla_ec_mg_m3_1", "vla_ec_mg_m3_2", "vla_ec_mg_m3_3", "vla_ec_mg_m3_4", "vla_ec_mg_m3_5", "vla_ec_mg_m3_6")
	Dim notas : notas = Array("notas_vla_1", "notas_vla_2", "notas_vla_3", "notas_vla_4", "notas_vla_5", "notas_vla_6")
	Dim i, valorLimiteAmbiental
	for i = 0 to ubound(estados)
			set valorLimiteAmbiental = extractValorLimiteAmbiental(substance, connection, _
				estados(i), _
				ed_ppm(i), _
				ed_mg_m3(i), _
				ec_ppm(i), _
				ec_mg_m3(i), _
				notas(i) _
			)
		if not isDictionaryEmpty(valorLimiteAmbiental) then
			result = arrayPushDictionary(result, valorLimiteAmbiental)
		end if
	next

	obtainValoresLimiteAmbiental = result
end function

function extractValorLimiteAmbiental(substance, connection, estado, ed_ppm, ed_mg_m3, ec_ppm, ec_mg_m3, notas)
	Dim result : set result = Server.CreateObject("Scripting.Dictionary")
	result.add "estado", substance.item(estado)
	result.add "ed_ppm", substance.item(ed_ppm)
	result.add "ed_mg_m3", substance.item(ed_mg_m3)
	result.add "ec_ppm", substance.item(ec_ppm)
	result.add "ec_mg_m3", substance.item(ec_mg_m3)
	result.add "notas", obtainNotasValoresLimiteAmbiental(substance.item(notas), connection)

	set extractValorLimiteAmbiental = result
end function

function isDictionaryEmpty(dictionary)
	isDictionaryEmpty = true
	if dictionary.count = 0 then exit function
	dim dictItems : dictItems = dictionary.items
	dim i, dictItem
	for i = 0 to ubound(dictItems)
		if hasValue(dictItems(i)) then 
			isDictionaryEmpty = false
			exit function
		end if
	next
end function

function hasValue(var)
	hasValue = false
	select case varType(var)
		case vbString
			if len(var) > 0 then hasValue = true
		case vbArray:
			if ubound(var) > -1 then hasValue = true
	end select
end function

function obtainNotasValoresLimiteAmbiental(byVal notasSrz, connection)
	dim result : result = Array()
	Dim notas
	if isNull(notasSrz) then 
		notasSrz = ""
	end if
	notas = split(notasSrz, ",")
	dim definitionKeysQueryFormatedSrz : definitionKeysQueryFormatedSrz = formatDefinitionKeysQueryList(notas)
	obtainNotasValoresLimiteAmbiental = findDefinitions(notas, definitionKeysQueryFormatedSrz, connection)
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