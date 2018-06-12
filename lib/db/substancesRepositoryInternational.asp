<!--#include file="synonymsRepository.asp"-->
<!--#include file="substanceListsRepository.asp"-->
<!--#include file="substanceGroupsRepositoryInternational.asp"-->
<!--#include file="substanceCompaniesRepository.asp"-->
<!--#include file="substanceApplicationsRepositoryInternational.asp"-->
<!--#include file="pictogramsRepository.asp"-->
<!--#include file="frasesRepository.asp"-->
<!--#include file="concentracionEtiquetadoRdRepository.asp"-->
<!--#include file="notasRdRepositoryInternational.asp"-->
<!--#include file="definitionsRepositoryInternational.asp"-->
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

function find_health_effects(substance_id, lang, connection)
	dim sql : sql = compose_health_effects_query(substance_id)
	dim recordset : set recordset = connection.execute(sql)
	dim health_effects : set health_effects = recodsetToDictionary(recordset)
	recordset.close()
	set recordset = nothing

	set find_health_effects = extract_health_effects(substance_id, lang, health_effects, connection)
end function

function find_environment_effects(substance_id, lang, connection)
	dim sql : sql = compose_environment_query(substance_id)
	dim recordset : set recordset = connection.execute(sql)
	dim environment_effects : set environment_effects = recodsetToDictionary(recordset)
	recordset.close()
	set recordset = nothing
	dim substance : set substance = extract_environment_effects(substance_id, lang, environment_effects, connection)

	set find_environment_effects = substance
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
	result.add "simbolos_rd1272", classification("simbolos_rd1272")
	result.Add "concentracionEtiquetadoRd1272", obtainConcentracionEtiquetadoRd1272(classification)

	result.add "clasificacion_rd1272_1", classification.item("clasificacion_rd1272_1")
	result.add "clasificacion_rd1272_2", classification.item("clasificacion_rd1272_2")
	result.add "clasificacion_rd1272_3", classification.item("clasificacion_rd1272_3")
	result.add "clasificacion_rd1272_4", classification.item("clasificacion_rd1272_4")
	result.add "clasificacion_rd1272_5", classification.item("clasificacion_rd1272_5")
	result.add "clasificacion_rd1272_6", classification.item("clasificacion_rd1272_6")
	result.add "clasificacion_rd1272_7", classification.item("clasificacion_rd1272_7")
	result.add "clasificacion_rd1272_8", classification.item("clasificacion_rd1272_8")
	result.add "clasificacion_rd1272_9", classification.item("clasificacion_rd1272_9")
	result.add "clasificacion_rd1272_10", classification.item("clasificacion_rd1272_10")
	result.add "clasificacion_rd1272_11", classification.item("clasificacion_rd1272_11")
	result.add "clasificacion_rd1272_12", classification.item("clasificacion_rd1272_12")
	result.add "clasificacion_rd1272_13", classification.item("clasificacion_rd1272_13")
	result.add "clasificacion_rd1272_14", classification.item("clasificacion_rd1272_14")
	result.add "clasificacion_rd1272_15", classification.item("clasificacion_rd1272_15")

	result.add "conc_rd1272_1", classification.item("conc_rd1272_1")
	result.add "conc_rd1272_2", classification.item("conc_rd1272_2")
	result.add "conc_rd1272_3", classification.item("conc_rd1272_3")
	result.add "conc_rd1272_4", classification.item("conc_rd1272_4")
	result.add "conc_rd1272_5", classification.item("conc_rd1272_5")
	result.add "conc_rd1272_6", classification.item("conc_rd1272_6")
	result.add "conc_rd1272_7", classification.item("conc_rd1272_7")
	result.add "conc_rd1272_8", classification.item("conc_rd1272_8")
	result.add "conc_rd1272_9", classification.item("conc_rd1272_9")
	result.add "conc_rd1272_10", classification.item("conc_rd1272_10")
	result.add "conc_rd1272_11", classification.item("conc_rd1272_11")
	result.add "conc_rd1272_12", classification.item("conc_rd1272_12")
	result.add "conc_rd1272_13", classification.item("conc_rd1272_13")
	result.add "conc_rd1272_14", classification.item("conc_rd1272_14")
	result.add "conc_rd1272_15", classification.item("conc_rd1272_15")

	result.add "eti_conc_rd1272_1", classification.item("eti_conc_rd1272_1")
	result.add "eti_conc_rd1272_2", classification.item("eti_conc_rd1272_2")
	result.add "eti_conc_rd1272_3", classification.item("eti_conc_rd1272_3")
	result.add "eti_conc_rd1272_4", classification.item("eti_conc_rd1272_4")
	result.add "eti_conc_rd1272_5", classification.item("eti_conc_rd1272_5")
	result.add "eti_conc_rd1272_6", classification.item("eti_conc_rd1272_6")
	result.add "eti_conc_rd1272_7", classification.item("eti_conc_rd1272_7")
	result.add "eti_conc_rd1272_8", classification.item("eti_conc_rd1272_8")
	result.add "eti_conc_rd1272_9", classification.item("eti_conc_rd1272_9")
	result.add "eti_conc_rd1272_10", classification.item("eti_conc_rd1272_10")
	result.add "eti_conc_rd1272_11", classification.item("eti_conc_rd1272_11")
	result.add "eti_conc_rd1272_12", classification.item("eti_conc_rd1272_12")
	result.add "eti_conc_rd1272_13", classification.item("eti_conc_rd1272_13")
	result.add "eti_conc_rd1272_14", classification.item("eti_conc_rd1272_14")
	result.add "eti_conc_rd1272_15", classification.item("eti_conc_rd1272_15")

	result.add "clasificacion_1", classification.item("clasificacion_1")
	result.add "clasificacion_2", classification.item("clasificacion_2")
	result.add "clasificacion_3", classification.item("clasificacion_3")
	result.add "clasificacion_4", classification.item("clasificacion_4")
	result.add "clasificacion_5", classification.item("clasificacion_5")
	result.add "clasificacion_6", classification.item("clasificacion_6")
	result.add "clasificacion_7", classification.item("clasificacion_7")
	result.add "clasificacion_8", classification.item("clasificacion_8")
	result.add "clasificacion_9", classification.item("clasificacion_9")
	result.add "clasificacion_10", classification.item("clasificacion_10")
	result.add "clasificacion_11", classification.item("clasificacion_11")
	result.add "clasificacion_12", classification.item("clasificacion_12")
	result.add "clasificacion_13", classification.item("clasificacion_13")
	result.add "clasificacion_14", classification.item("clasificacion_14")
	result.add "clasificacion_15", classification.item("clasificacion_15")

	result.Add "notas_rd1272", obtainNotasRd1272(classification.item("notas_rd1272"), lang, connection)

	set extractClassification = result
End Function

function extract_health_effects(substanceId, lang, substanceDic, connection)
	dim substance : set substance = Server.CreateObject("Scripting.Dictionary")
	dim featuredLists : featuredLists = obtainFeaturedLists(substanceId, connection)
	dim substanceGroupsRecordset
	set substanceGroupsRecordset = getRecordsetSubstanceGroupsInternational(substanceId, lang, connection)
	set substanceDic = addSubstanceGroupsAssociatedFields(substanceDic, substanceGroupsRecordset)
	substanceGroupsRecordset.close()
	set substanceGroupsRecordset = nothing
	substance.add "comentarios_sl", substanceDic("comentarios_sl")
	substance.add "grupo_iarc", extractGrupoIarc(substanceDic("grupo_iarc"))
	substance.add "volumen_iarc", substanceDic("volumen_iarc")
	substance.add "notas_iarc", substanceDic("notas_iarc")
	substance.add "nivel_disruptor", obtainDefinitions(substanceDic("nivel_disruptor"), lang, connection)
	substance.add "efecto_neurotoxico", obtainEfectosNeurotoxico(substanceDic("efecto_neurotoxico"), featuredLists, lang, connection)
	substance.add "fuente_neurotoxico", obtainFuentesNeurotoxico(substanceDic("fuente_neurotoxico"), featuredLists, lang, connection)
	dim nivel_neurotoxico_key
	nivel_neurotoxico_key = obtainNivelNeurotoxicoKey(substanceDic("nivel_neurotoxico"))
	substance.add "nivel_neurotoxico", obtainDefinitions(nivel_neurotoxico_key, lang, connection)
	substance.add "nivel_tpr", obtainNivelTpr(substanceDic, connection)
	substance.add "categoria_cancer_otras", substanceDic("categoria_cancer_otras")
	substance.add "fuente", substanceDic("fuente")

	set extract_health_effects = substance
end function

function extract_environment_effects(substance_id, lang, substanceDic, connection)
	dim substance : set substance = Server.CreateObject("Scripting.Dictionary")
	dim substanceGroupsRecordset : set substanceGroupsRecordset = getRecordsetSubstanceGroupsInternational(substance_id, lang, connection)
	set substanceDic = addSubstanceGroupsAssociatedFields(substanceDic, substanceGroupsRecordset)
	substanceGroupsRecordset.close()
	set substanceGroupsRecordset = nothing
	substance.add "anchor_tpb", substanceDic("anchor_tpb")
	substance.add "enlace_tpb", substanceDic("enlace_tpb")
	substance.add "fuentes_tpb", obtainDefinitions(substanceDic("fuentes_tpb"), lang, connection)
	substance.add "directiva_aguas", substanceDic("directiva_aguas")
	substance.add "clasif_mma", obtainClasifMma(substanceDic("clasif_mma"), lang, connection)
	substance.add "sustancia_prioritaria", substanceDic("sustancia_prioritaria")

	set extract_environment_effects = substance
end function

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

function compose_health_effects_query(id_sustancia)
	dim sql
sql = _
		"SELECT " &_
			"sus.id, sus.comentarios as comentarios_sl, iarc.grupo_iarc, iarc.notas_iarc, iarc.volumen_iarc, " &_
			"neurodis.nivel_disruptor, neurodis.efecto_neurotoxico, neurodis.fuente_neurotoxico, neurodis.nivel_neurotoxico, cancer_otras.categoria_cancer_otras, cancer_otras.fuente, " &_
			"sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, " &_
			"sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, " &_
			"sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15 " &_
		"FROM " &_
			"dn_risc_sustancias as sus " &_
		"LEFT JOIN " &_
			"dn_risc_sustancias_iarc as iarc " &_
				"ON sus.id = iarc.id_sustancia " &_
		"LEFT JOIN " &_
			"dn_risc_sustancias_neuro_disruptor as neurodis " &_
				"ON sus.id = neurodis.id_sustancia " &_
		"LEFT JOIN " &_
			"dn_risc_sustancias_cancer_otras as cancer_otras " &_
				"ON sus.id = cancer_otras.id_sustancia " &_
		"WHERE " &_
			"sus.id = " & id_sustancia

	compose_health_effects_query = sql
end function

function compose_environment_query(substanceId)
	dim sql
	sql = _
		"SELECT " &_
			"anchor_tpb, enlace_tpb, fuentes_tpb, mpmb, " &_
			"directiva_aguas, clasif_mma, sustancia_prioritaria " &_
		"FROM " &_
			"dn_risc_sustancias_ambiente " &_
		"WHERE " &_
			"id_sustancia = " & substanceId

	compose_environment_query = sql
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

function extractGrupoIarc(grupo)
	extractGrupoIarc = grupo
	if isNull(grupo) then
		exit function
	end if
	extractGrupoIarc = ""
	extractGrupoIarc = _
		trim( _
			replace( _
				ucase(grupo), "GRUPO", "") _
			)
end function

function obtainEfectosNeurotoxico(byVal efectosSrz, featuredLists, lang, connection)
	obtainEfectosNeurotoxico = efectosSrz
	if isNull(efectosSrz) then _
		exit function
	obtainEfectosNeurotoxico = obtainDefinitions( _
		replace(efectosSrz, "/", ",") _
		, lang, connection)
	dim efectos : efectos = split(efectosSrz, "/")
	if not( _
		inArray("neurotoxico_rd", featuredLists) _
		or inArray("neurotoxico_danesa", featuredLists) _
		) then exit function
	if not inArray("SNC", efectos) then _
		arrayPush efectos, "SNC"

	obtainEfectosNeurotoxico = obtainDefinitions( _
		join(efectos, ",") _
		, lang, connection )
end function

function obtainFuentesNeurotoxico(fuentesSrz, featuredLists, lang, connection)
	obtainFuentesNeurotoxico = obtainDefinitions(fuentesSrz, lang, connection)
	if isNull(fuentesSrz) then _
		exit function
	dim fuentes : fuentes = split(fuentesSrz)
	if not( _
		inArray("neurotoxico_rd", featuredLists) _
		or inArray("neurotoxico_danesa", featuredLists) _
		) then exit function
	if not inArray("363", fuentes) then _
		arrayPush fuentes, "363"
	if not inArray("EPA-R67", featuredLists) then _
		arrayPush fuentes, "EPA-R67"

	obtainFuentesNeurotoxico = obtainDefinitions( _
		join(fuentes, ",") _
		, lang, connection )
end function

function obtainNivelNeurotoxicoKey(nivel)
	obtainNivelNeurotoxicoKey = nivel
	if isNull(nivel) or nivel = "" then _
		exit function

	obtainNivelNeurotoxicoKey = "Nivel " &  nivel
end function

function obtainNivelTpr(substanceDic, connection)
	set obtainNivelTpr = Server.CreateObject("Scripting.Dictionary")
	dim clasificaciones : clasificaciones = _
		substanceDic("clasificacion_1") &_
		substanceDic("clasificacion_2") &_
		substanceDic("clasificacion_3") &_
		substanceDic("clasificacion_4") &_
		substanceDic("clasificacion_5") &_
		substanceDic("clasificacion_6") &_
		substanceDic("clasificacion_7") &_
		substanceDic("clasificacion_8") &_
		substanceDic("clasificacion_9") &_
		substanceDic("clasificacion_10") &_
		substanceDic("clasificacion_11") &_
		substanceDic("clasificacion_12") &_
		substanceDic("clasificacion_13") &_
		substanceDic("clasificacion_14") &_
		substanceDic("clasificacion_15")
	clasificaciones = replace(clasificaciones, " ", "")
	dim posicion : posicion = instr(clasificaciones, "Repr.Cat.")
	if posicion = 0 then _
		exit function

	obtainNivelTpr = obtainDefinitions( "TR" &_
		replace( _
			replace( _
				replace( _
					mid(clasificaciones, posicion + 9, 1), "1", "1A" _
				), "2", "1B"_
			), "3", "2" _
		) _
	, lang, connection)
end function

function obtainClasifMma(byVal clasif, lang, connection)
	dim key : key = ""
	obtainClasifMma = ""
	if not isNumeric(clasif) then _
		exit function
	if (0 < clasif < 4 ) then _
		key = clasif + "."

	obtainClasifMma = obtainDefinitions(key, lang, connection)
end function
%>
