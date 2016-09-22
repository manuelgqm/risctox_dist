<%
function obtainNotasRd1272(notas, connection)
	Dim result : result = Array()
	if varType(notas) = vbNull then
		obtainNotasRd1272 = result
		Exit function
	end if
	if notas = "" then
		obtainNotasRd1272 = result
		Exit function
	end if 
	notas = replaceValidated(notas, "@", "@ ")
	notas = replaceValidated(notas, ",", "")
	notas = removeTailSeparator(notas, ".")
	Dim notasList : notasList = split(notas, ".")
	Dim i, notaValue, notaId
	Dim nota
	Dim notasDefinitions : set notasDefinitions = findNotasDefinition(notasList, connection)
	for i = 0 to ubound(notasList)
		set nota = Server.CreateObject("Scripting.Dictionary")
		notaValue = notasList(i)
		nota.add "value", notaValue
		nota.add "id", notasDefinitions.item(notaValue).item("id")
		nota.add "description", notasDefinitions.item(notaValue).item("description")
		result = arrayPushDictionary(result, nota)
		set nota = nothing
	next
	
	obtainNotasRd1272 = result
end function

function findNotasDefinition(notasList, connection)
	Dim notasDefinitionsQuery : notasDefinitionsQuery = composeNotasDefinitionsQuery(notasList)
	Dim notasDefinitionsRecordSet : set notasDefinitionsRecordSet = connection.execute(notasDefinitionsQuery)
	Dim result : set result = Server.CreateObject("Scripting.Dictionary")
	Dim i, notaId, notaDescription
	For i = 0 to Ubound(notasList)
		set nota = Server.CreateObject("Scripting.Dictionary")
		notaId = ""
		notaDescription = ""
		if not notasDefinitionsRecordSet.EOF then
			notaId = notasDefinitionsRecordSet("id").value
			notaDescription = notasDefinitionsRecordSet("definicion").value
			notasDefinitionsRecordSet.moveNext
		end if
		nota.add "id", notaId
		nota.add "description", notaDescription
		result.add notasList(i), nota
	next

	notasDefinitionsRecordSet.close
	set notasDefinitionsRecordSet = nothing
	set findNotasDefinition = result
end function

function composeNotasDefinitionsQuery(notasList)
	Dim notasListQueryFormated : notasListQueryFormated = arrayWrapItems(notasList, "'R.1272-", "'")
	Dim notasListQueryFormatedSerialized : notasListQueryFormatedSerialized = join(notasListQueryFormated, ",")
	Dim sql
	sql = "SELECT id, palabra, dbo.udf_StripHTML(definicion) as definicion FROM rq_definiciones where palabra in (" & notasListQueryFormatedSerialized & ")"
	composeNotasDefinitionsQuery = sql
end function
%>