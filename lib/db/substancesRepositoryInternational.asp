<!--#include file="synonymsRepository.asp"-->
<!--#include file="substanceListsRepository.asp"-->
<!--#include file="substanceGroupsRepositoryInternational.asp"-->
<!--#include file="substanceApplicationsRepositoryInternational.asp"-->
<!--#include file="pictogramsRepository.asp"-->
<!--#include file="frasesRepository.asp"-->
<!--#include file="concentracionEtiquetadoRdRepository.asp"-->
<!--#include file="substanceCompaniesRepository.asp"-->
<%
function findIdentification(substance_id, lang, connection)
	dim sql_query : sql_query = composeIdentificationQuery(substance_id)
	dim identification_rs : set identification_rs = connection.execute(sql_query)
	dim identification : set identification = recodsetToDictionary(identification_rs)
	identification_rs.close()
	set identification_rs = nothing

	set findIdentification = extractIdentification(substance_id, identification, lang, connection)
end function

Function findClassification(substance_id, lang, connection)
	dim sql_query : sql_query = composeClassificationQuery(substance_id)
	dim recordset : set recordset = connection.execute(sql_query)
	dim result : set result = recodsetToDictionary(recordset)
	recordset.close()
	set recordset = nothing

	set findClassification = extractClassification(substance_id, result, lang, connection)
End Function

' PRIVATE
function extractIdentification(substance_id, identification, lang, connection)
	dim result : set result = Server.CreateObject("Scripting.Dictionary")
	result.add "nombre", obtainNombre(identification("nombre"), lang)
	result.add "nombre_ing", identification("nombre_ing")
	' result.add "sinonimos", obtainSynonyms(substanceId, connection)
	result.add "num_cas", identification("num_cas")
	result.add "num_ce_einecs", identification("num_ce_einecs")
	result.add "num_ce_elincs", identification("num_ce_elincs")
	result.add "num_rd", identification("num_rd")
	dim substanceGroupsRecordset
	set substanceGroupsRecordset = getRecordsetSubstanceGroupsInternational(substance_id, lang, connection)
	result.Add "grupos", extractSubstanceGroups(substanceGroupsRecordset)
	set result = addSubstanceGroupsAssociatedFields(result, substanceGroupsRecordset)
	substanceGroupsRecordset.close()
	set substanceGroupsRecordset = nothing
	result.add "applications", findSubstanceApplicationsInternational(substance_id, lang, connection)
	result.add "icsc_nums", obtainNumsIcsc(identification("num_icsc"))
	result.add "molecular_formula", identification("formula_molecular")
	result.add "molecular_structure", identification("estructura_molecular")
	result.Add "featuredLists", obtainFeaturedLists(id_sustancia, connection)

	set extractIdentification = result
end function

Function extractClassification(substance_id, classification, lang, connection)
	dim result : set result = Server.CreateObject("Scripting.Dictionary")
	result.Add "frasesR", joinFrases("R", classification)
	result.Add "frasesH", findFrasesH(classification, connection)
	result.Add "pictogramasRd1272", findPictograms(classification("simbolos_rd1272"), connection)
	result.Add "concentracionEtiquetadoRd1272", obtainConcentracionEtiquetadoRd1272(classification)

	set extractClassification = result
End Function

Function composeClassificationQuery(substance_id)
	composeClassificationQuery = _
		"SELECT " &_
			"clasificacion_1, clasificacion_2, clasificacion_3, " &_
			"clasificacion_4, clasificacion_5, clasificacion_6, " &_
			"clasificacion_7, clasificacion_8, clasificacion_9, " &_
			"clasificacion_10, clasificacion_11, clasificacion_12, " &_
			"clasificacion_13, clasificacion_14, clasificacion_15, " &_
			"clasificacion_rd1272_1, clasificacion_rd1272_2, clasificacion_rd1272_3, " &_
			"clasificacion_rd1272_4, clasificacion_rd1272_5, clasificacion_rd1272_6, " &_
			"clasificacion_rd1272_7, clasificacion_rd1272_8, clasificacion_rd1272_9, " &_
			"clasificacion_rd1272_10, clasificacion_rd1272_11, clasificacion_rd1272_12, " &_
			"clasificacion_rd1272_13, clasificacion_rd1272_14, clasificacion_rd1272_15, " &_
			"conc_rd1272_1, conc_rd1272_2, conc_rd1272_3, conc_rd1272_4, conc_rd1272_5, " &_
			"conc_rd1272_6, conc_rd1272_7, conc_rd1272_8, conc_rd1272_9, conc_rd1272_10, " &_
			"conc_rd1272_11, conc_rd1272_12, conc_rd1272_13, conc_rd1272_14, conc_rd1272_15, " &_
			"eti_conc_rd1272_1, eti_conc_rd1272_2, eti_conc_rd1272_3, " &_
			"eti_conc_rd1272_4, eti_conc_rd1272_5, eti_conc_rd1272_6, " &_
			"eti_conc_rd1272_7, eti_conc_rd1272_8, eti_conc_rd1272_9, " &_
			"eti_conc_rd1272_10, eti_conc_rd1272_11, eti_conc_rd1272_12, " &_
			"eti_conc_rd1272_13, eti_conc_rd1272_14, eti_conc_rd1272_15, " &_
			"notas_rd1272, simbolos_rd1272, clases_categorias_peligro_rd1272, " &_
			"frases_r_danesa " &_
		"FROM " &_
			"dn_risc_sustancias " &_
		"WHERE " &_
			"id = " & substance_id
End Function

function composeIdentificationQuery(substance_id)
	composeIdentificationQuery = _
		"SELECT " &_
			"nombre_ing as nombre, nombre_ing, num_rd, num_ce_einecs, num_ce_elincs, num_cas, " &_
			"cas_alternativos, num_icsc, formula_molecular, estructura_molecular " &_
		"FROM " &_
			"dn_risc_sustancias " &_
		"WHERE " &_
			"id = " & substance_id
end function

function obtainNombre(nombre, lang)
	if lang = "en" then
		dim names : names =	split(nombre, "@")
		obtainNombre = names(0)
		exit function
	end if

	obtainNombre = nombre
end function

function obtainNumsIcsc(numsIcscSrz)
	dim icsc
	dim result : result = Array()
	dim numsIcsc : numsIcsc = split(numsIcscSrz, "@")
	dim i, centena, max, min
	for i = 0 to ubound(numsIcsc)
		current = cstr(numsIcsc(i))
		if len(current) <> 4 then
			obtainNumsIcsc = result
			exit function
		end if
		centena = mid(current, 1, 2)
		max = cstr(clng(centena & "01"))
		if max = "1" then max = "0"
		min = cstr(clng(centena) + 1) & "00"
		set icsc = Server.CreateObject("Scripting.Dictionary")
		icsc.add "id", current
		icsc.add "max", max
		icsc.add "min", min
		result = arrayPushDictionary(result, icsc)
	next

	obtainNumsIcsc = result
end function

sub printSusbtance(substance)
	for each key in substance.keys
		response.write key & ": "
		if isArray(substance.item(key)) then
			for k = 0 to ubound(substance.item(key))
				if vartype(substance.item(key)(k)) = 9 then
					for each u in substance.item(key)(k)
						response.write substance.item(key)(k).item(u)
					next
				else
					response.write substance.item(key)(k) & ","
				end if
			next
		else
			response.write substance.item(key)
		end if
		response.write "<br>"
	next
end sub

function recodsetToDictionary(recordset)
	set result = Server.CreateObject("Scripting.Dictionary")
	if recordset.eof then
		set recodsetToDictionary = result
		exit function
	end if
	dim key
	for each key in recordset.fields
		result.add key.name, key.Value
	next
	set recodsetToDictionary = result
end function

function recodsetToDictionaryArray(recordset)
	dim result : result = Array()
	if recordset.eof then
		recodsetToDictionaryArray = result
		exit function
	end if
	while not recordset.eof
		result = arrayPushDictionary(result, recodsetToDictionary(recordset))
		recordset.movenext
	wend

	recodsetToDictionaryArray = result
end function

function joinFrases(tipo, classification)
	frases = ""
	frases = extractFrase(classification("clasificacion_1"), frases, tipo)
	frases = extractFrase(classification("clasificacion_2"), frases, tipo)
	frases = extractFrase(classification("clasificacion_3"), frases, tipo)
	frases = extractFrase(classification("clasificacion_4"), frases, tipo)
	frases = extractFrase(classification("clasificacion_5"), frases, tipo)
	frases = extractFrase(classification("clasificacion_6"), frases, tipo)
	frases = extractFrase(classification("clasificacion_7"), frases, tipo)
	frases = extractFrase(classification("clasificacion_8"), frases, tipo)
	frases = extractFrase(classification("clasificacion_9"), frases, tipo)
	frases = extractFrase(classification("clasificacion_10"), frases, tipo)
	frases = extractFrase(classification("clasificacion_11"), frases, tipo)
	frases = extractFrase(classification("clasificacion_12"), frases, tipo)
	frases = extractFrase(classification("clasificacion_13"), frases, tipo)
	frases = extractFrase(classification("clasificacion_14"), frases, tipo)
	frases = extractFrase(classification("clasificacion_15"), frases, tipo)

	joinFrases=frases
end function

function extractFrase(c,f, tipo)
	c = formatFrases(c, tipo)

	' Limpiamos la clasificación para quedarnos con las frases
	array_frases = split(c, " ")
	nuevo_c = ""
	for i=0 to ubound(array_frases)
		' Para que sea frase R ha de tener longitud 2 o mayor y comenzar por R más un digito ( o H)
		' Ej.: "R1", "R10", "R1/6"

		if (es_frase(array_frases(i),tipo)) then
			if (nuevo_c="") then
				nuevo_c = array_frases(i)
			else
				nuevo_c = nuevo_c&", "&array_frases(i)
			end if
		end if
	next

	if (nuevo_c <> "") then
		' La clasificación no es vacía, concatenamos a la frase
		if (f <> "") then
			' Ya hay algo en las frases, concateno
			extractFrase = f & ", " & nuevo_c
		else
			' No hay nada, devuelvo clasificación
			extractFrase = nuevo_c
		end if
	else
		' La clasificacion es vacía, devolvemos la frase tal cual
		extractFrase = f
	end if
end function

function formatFrases(byval c, tipo)
	' En casos como el DDT que tiene frases como R25-48/25, hay que convertir a R25 R48/25, o sea, cambiar "-" por " R",
	' pero solo en los casos en que tenga "-" y "/"

	' Lo dividimos primero separando por espacios, arreglamos cada una y lo volvemos a unir
	c2=""
	if isnull(c) then c = ""
	array_c = split(c, " ")
	for i=0 to ubound(array_c)
		if ((instr(array_c(i), "-") <> 0) and (instr(array_c(i), "/") <> 0)) then
			array_c(i)=replace(array_c(i), "-", " "+tipo)
		end if
		c2=c2&" "&array_c(i)
	next

'	response.write "<br />"&c&" se convierte en "&c2

	formatFrases=c2
end function
%>
