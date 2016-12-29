<!--#include file="synonymsRepository.asp"-->
<!--#include file="substanceListsRepository.asp"-->
<!--#include file="substanceGroupsRepository.asp"-->
<!--#include file="substanceApplicationsRepository.asp"-->
<!--#include file="substanceCompaniesRepository.asp"-->
<!--#include file="pictogramsRepository.asp"-->
<!--#include file="classificationsRd1272Repository.asp"-->
<!--#include file="notasRd1272Repository.asp"-->
<!--#include file="concentracionEtiquetadoRd1272Repository.asp"-->
<!--#include file="valoresLimiteAmbientalRepository.asp"-->
<!--#include file="valoresLimiteBiologicoRepository.asp"-->
<!--#include file="definitionsRepository.asp"-->
<%
function findSubstance(id_sustancia, connection)
	sql = composeSubstanceQuery( id_sustancia )
	set substanceRecordset = connection.execute(sql)
	set substance = extractSubstance(id_sustancia, substanceRecordset, connection)
	substanceRecordset.close()
	set substanceRecordset=nothing
	set findSubstance = substance
end function

function findSubstanceLevelOne(id_sustancia, connection)
	sql = composeSubstanceLevelOneFieldsQuery( id_sustancia )
	set substanceRecordset = connection.execute(sql)
	set substanceDic = recodsetToDictionary(substanceRecordset)
	substanceRecordset.close()
	set substance = extractSubstanceLevelOneFields(id_sustancia, substanceDic, connection)
	set substanceRecordset = nothing
	set findSubstanceLevelOne = substance
end function

function findSaludFields(id_sustancia, connection)
	dim sql : sql = composeSaludQuery(id_sustancia)
	dim substanceRecordset : set substanceRecordset = connection.execute(sql)
	dim substanceDic : set substanceDic = recodsetToDictionary(substanceRecordset)
	substanceRecordset.close()
	set substanceRecordset = nothing
	dim substance : set substance = extractSubstanceSaludFields(id_sustancia, substanceDic, connection)

	set findSaludFields = substance
end function

function findMedioAmbienteFields(id_sustancia, connection)
	dim sql : sql = composeMedioAmbienteQuery(id_sustancia)
	dim substanceRecordset : set substanceRecordset = connection.execute(sql)
	dim substanceDic : set substanceDic = recodsetToDictionary(substanceRecordset)
	substanceRecordset.close()
	set substanceRecordset = nothing
	dim substance : set substance = extractSubstanceMedioAmbienteFields(id_sustancia, substanceDic, connection)

	set findMedioAmbienteFields = substance
end function

function findCancerOtrasFields(id_sustancia, connection)
	dim sql : sql = composeCancerOtrasQuery(id_sustancia)
	dim substanceRecordset : set substanceRecordset = connection.execute(sql)
	dim substanceDic : set substanceDic = recodsetToDictionary(substanceRecordset)
	substanceRecordset.close()
	set substanceRecordset = nothing
	dim substance : set substance = extractCancerOtrasFields(id_sustancia, substanceDic, connection)

	set findCancerOtrasFields = substance
end function

function findEnfermedadesFields(id_sustancia, connection)
	dim sql : sql = composeEnfermedadesQuery(id_sustancia)
	dim substanceRecordset : set substanceRecordset = connection.execute(sql)
	dim substanceDics : substanceDics = recodsetToDictionaryArray(substanceRecordset)
	substanceRecordset.close()
	set substanceRecordset = nothing
	dim substance : set substance = extractEnfermedadesFields(substanceDics)
	
	set findEnfermedadesFields = substance
end function

' PRIVATE
function extractSubstance(id_sustancia, substanceRecordset, connection)
	set substance = Server.CreateObject("Scripting.Dictionary")

	' dn_risc_sustancias
	substance.Add "nombre", substanceRecordset("nombre").Value
	substance.Add "nombre_ing", removeDuplicates(substanceRecordset("nombre_ing").Value, "@")

	substance.Add "num_rd", substanceRecordset("num_rd").Value
	substance.Add "num_ce_einecs", substanceRecordset("num_ce_einecs").Value
	substance.Add "num_ce_elincs", substanceRecordset("num_ce_elincs").Value
	substance.Add "num_cas", substanceRecordset("num_cas").Value
	substance.Add "cas_alternativos", substanceRecordset("cas_alternativos").Value

	substance.Add "num_icsc", substanceRecordset("num_icsc").Value
	substance.Add "formula_molecular", substanceRecordset("formula_molecular").Value
	substance.Add "estructura_molecular", substanceRecordset("estructura_molecular").Value
	substance.Add "simbolos", substanceRecordset("simbolos").Value
	substance.Add "clasificacion_1", trim(substanceRecordset("clasificacion_1").Value)
	substance.Add "clasificacion_2", trim(substanceRecordset("clasificacion_2").Value)
	substance.Add "clasificacion_3", trim(substanceRecordset("clasificacion_3").Value)
	substance.Add "clasificacion_4", trim(substanceRecordset("clasificacion_4").Value)
	substance.Add "clasificacion_5", trim(substanceRecordset("clasificacion_5").Value)
	substance.Add "clasificacion_6", trim(substanceRecordset("clasificacion_6").Value)
	substance.Add "clasificacion_7", trim(substanceRecordset("clasificacion_7").Value)
	substance.Add "clasificacion_8", trim(substanceRecordset("clasificacion_8").Value)
	substance.Add "clasificacion_9", trim(substanceRecordset("clasificacion_9").Value)
	substance.Add "clasificacion_10", trim(substanceRecordset("clasificacion_10").Value)
	substance.Add "clasificacion_11", trim(substanceRecordset("clasificacion_11").Value)
	substance.Add "clasificacion_12", trim(substanceRecordset("clasificacion_12").Value)
	substance.Add "clasificacion_13", trim(substanceRecordset("clasificacion_13").Value)
	substance.Add "clasificacion_14", trim(substanceRecordset("clasificacion_14").Value)
	substance.Add "clasificacion_15", trim(substanceRecordset("clasificacion_15").Value)
	substance.Add "frases_s", trim(substanceRecordset("frases_s").Value)
	substance.Add "conc_1", substanceRecordset("conc_1").Value
	substance.Add "eti_conc_1", substanceRecordset("eti_conc_1").Value
	substance.Add "conc_2", substanceRecordset("conc_2").Value
	substance.Add "eti_conc_2", substanceRecordset("eti_conc_2").Value
	substance.Add "conc_3", substanceRecordset("conc_3").Value
	substance.Add "eti_conc_3", substanceRecordset("eti_conc_3").Value
	substance.Add "conc_4", substanceRecordset("conc_4").Value
	substance.Add "eti_conc_4", substanceRecordset("eti_conc_4").Value
	substance.Add "conc_5", substanceRecordset("conc_5").Value
	substance.Add "eti_conc_5", substanceRecordset("eti_conc_5").Value
	substance.Add "conc_6", substanceRecordset("conc_6").Value
	substance.Add "eti_conc_6", substanceRecordset("eti_conc_6").Value
	substance.Add "conc_7", substanceRecordset("conc_7").Value
	substance.Add "eti_conc_7", substanceRecordset("eti_conc_7").Value
	substance.Add "conc_8", substanceRecordset("conc_8").Value
	substance.Add "eti_conc_8", substanceRecordset("eti_conc_8").Value
	substance.Add "conc_9", substanceRecordset("conc_9").Value
	substance.Add "eti_conc_9", substanceRecordset("eti_conc_9").Value
	substance.Add "conc_10", substanceRecordset("conc_10").Value
	substance.Add "eti_conc_10", substanceRecordset("eti_conc_10").Value
	substance.Add "conc_11", substanceRecordset("conc_11").Value
	substance.Add "eti_conc_11", substanceRecordset("eti_conc_11").Value
	substance.Add "conc_12", substanceRecordset("conc_12").Value
	substance.Add "eti_conc_12", substanceRecordset("eti_conc_12").Value
	substance.Add "conc_13", substanceRecordset("conc_13").Value
	substance.Add "eti_conc_13", substanceRecordset("eti_conc_13").Value
	substance.Add "conc_14", substanceRecordset("conc_14").Value
	substance.Add "eti_conc_14", substanceRecordset("eti_conc_14").Value
	substance.Add "conc_15", substanceRecordset("conc_15").Value
	substance.Add "eti_conc_15", substanceRecordset("eti_conc_15").Value
	substance.Add "notas_rd_363", substanceRecordset("notas_rd_363").Value
	substance.Add "notas_xml", replaceValidated(substanceRecordset("notas_xml").Value, "@", "@ ")
	substance.Add "frases_r_danesa", trim(substanceRecordset("frases_r_danesa").Value)

	' RD1272/2008
	substance.Add "clasificacion_rd1272_1", trim(substanceRecordset("clasificacion_rd1272_1").Value)
	substance.Add "clasificacion_rd1272_2", trim(substanceRecordset("clasificacion_rd1272_2").Value)
	substance.Add "clasificacion_rd1272_3", trim(substanceRecordset("clasificacion_rd1272_3").Value)
	substance.Add "clasificacion_rd1272_4", trim(substanceRecordset("clasificacion_rd1272_4").Value)
	substance.Add "clasificacion_rd1272_5", trim(substanceRecordset("clasificacion_rd1272_5").Value)
	substance.Add "clasificacion_rd1272_6", trim(substanceRecordset("clasificacion_rd1272_6").Value)
	substance.Add "clasificacion_rd1272_7", trim(substanceRecordset("clasificacion_rd1272_7").Value)
	substance.Add "clasificacion_rd1272_8", trim(substanceRecordset("clasificacion_rd1272_8").Value)
	substance.Add "clasificacion_rd1272_9", trim(substanceRecordset("clasificacion_rd1272_9").Value)
	substance.Add "clasificacion_rd1272_10", trim(substanceRecordset("clasificacion_rd1272_10").Value)
	substance.Add "clasificacion_rd1272_11", trim(substanceRecordset("clasificacion_rd1272_11").Value)
	substance.Add "clasificacion_rd1272_12", trim(substanceRecordset("clasificacion_rd1272_12").Value)
	substance.Add "clasificacion_rd1272_13", trim(substanceRecordset("clasificacion_rd1272_13").Value)
	substance.Add "clasificacion_rd1272_14", trim(substanceRecordset("clasificacion_rd1272_14").Value)
	substance.Add "clasificacion_rd1272_15", trim(substanceRecordset("clasificacion_rd1272_15").Value)
	substance.Add "conc_rd1272_1", substanceRecordset("conc_rd1272_1").Value
	substance.Add "eti_conc_rd1272_1", substanceRecordset("eti_conc_rd1272_1").Value
	substance.Add "conc_rd1272_2", substanceRecordset("conc_rd1272_2").Value
	substance.Add "eti_conc_rd1272_2", substanceRecordset("eti_conc_rd1272_2").Value
	substance.Add "conc_rd1272_3", substanceRecordset("conc_rd1272_3").Value
	substance.Add "eti_conc_rd1272_3", substanceRecordset("eti_conc_rd1272_3").Value
	substance.Add "conc_rd1272_4", substanceRecordset("conc_rd1272_4").Value
	substance.Add "eti_conc_rd1272_4", substanceRecordset("eti_conc_rd1272_4").Value
	substance.Add "conc_rd1272_5", substanceRecordset("conc_rd1272_5").Value
	substance.Add "eti_conc_rd1272_5", substanceRecordset("eti_conc_rd1272_5").Value
	substance.Add "conc_rd1272_6", substanceRecordset("conc_rd1272_6").Value
	substance.Add "eti_conc_rd1272_6", substanceRecordset("eti_conc_rd1272_6").Value
	substance.Add "conc_rd1272_7", substanceRecordset("conc_rd1272_7").Value
	substance.Add "eti_conc_rd1272_7", substanceRecordset("eti_conc_rd1272_7").Value
	substance.Add "conc_rd1272_8", substanceRecordset("conc_rd1272_8").Value
	substance.Add "eti_conc_rd1272_8", substanceRecordset("eti_conc_rd1272_8").Value
	substance.Add "conc_rd1272_9", substanceRecordset("conc_rd1272_9").Value
	substance.Add "eti_conc_rd1272_9", substanceRecordset("eti_conc_rd1272_9").Value
	substance.Add "conc_rd1272_10", substanceRecordset("conc_rd1272_10").Value
	substance.Add "eti_conc_rd1272_10", substanceRecordset("eti_conc_rd1272_10").Value
	substance.Add "conc_rd1272_11", substanceRecordset("conc_rd1272_11").Value
	substance.Add "eti_conc_rd1272_11", substanceRecordset("eti_conc_rd1272_11").Value
	substance.Add "conc_rd1272_12", substanceRecordset("conc_rd1272_12").Value
	substance.Add "eti_conc_rd1272_12", substanceRecordset("eti_conc_rd1272_12").Value
	substance.Add "conc_rd1272_13", substanceRecordset("conc_rd1272_13").Value
	substance.Add "eti_conc_rd1272_13", substanceRecordset("eti_conc_rd1272_13").Value
	substance.Add "conc_rd1272_14", substanceRecordset("conc_rd1272_14").Value
	substance.Add "eti_conc_rd1272_14", substanceRecordset("eti_conc_rd1272_14").Value
	substance.Add "conc_rd1272_15", substanceRecordset("conc_rd1272_15").Value
	substance.Add "eti_conc_rd1272_15", substanceRecordset("eti_conc_rd1272_15").Value
	substance.Add "notas_rd1272", obtainNotasRd1272(substanceRecordset("notas_rd1272"), connection)
	substance.Add "simbolos_rd1272", substanceRecordset("simbolos_rd1272").Value
	substance.Add "clases_categorias_peligro_rd1272", substanceRecordset("clases_categorias_peligro_rd1272").Value

	' dn_risc_sustancias_vl
	substance.Add "estado_1", substanceRecordset("estado_1").Value
	substance.Add "vla_ed_ppm_1", substanceRecordset("vla_ed_ppm_1").Value
	substance.Add "vla_ed_mg_m3_1", substanceRecordset("vla_ed_mg_m3_1").Value
	substance.Add "vla_ec_ppm_1", substanceRecordset("vla_ec_ppm_1").Value
	substance.Add "vla_ec_mg_m3_1", substanceRecordset("vla_ec_mg_m3_1").Value
	substance.Add "notas_vla_1", removeVlbFromNotes(substanceRecordset("notas_vla_1").Value)

	substance.Add "estado_2", substanceRecordset("estado_2").Value
	substance.Add "vla_ed_ppm_2", substanceRecordset("vla_ed_ppm_2").Value
	substance.Add "vla_ed_mg_m3_2", substanceRecordset("vla_ed_mg_m3_2").Value
	substance.Add "vla_ec_ppm_2", substanceRecordset("vla_ec_ppm_2").Value
	substance.Add "vla_ec_mg_m3_2", substanceRecordset("vla_ec_mg_m3_2").Value
	substance.Add "notas_vla_2", removeVlbFromNotes(substanceRecordset("notas_vla_2").Value)

	substance.Add "estado_3", substanceRecordset("estado_3").Value
	substance.Add "vla_ed_ppm_3", substanceRecordset("vla_ed_ppm_3").Value
	substance.Add "vla_ed_mg_m3_3", substanceRecordset("vla_ed_mg_m3_3").Value
	substance.Add "vla_ec_ppm_3", substanceRecordset("vla_ec_ppm_3").Value
	substance.Add "vla_ec_mg_m3_3", substanceRecordset("vla_ec_mg_m3_3").Value
	substance.Add "notas_vla_3", removeVlbFromNotes(substanceRecordset("notas_vla_3").Value)

	substance.Add "estado_4", substanceRecordset("estado_4").Value
	substance.Add "vla_ed_ppm_4", substanceRecordset("vla_ed_ppm_4").Value
	substance.Add "vla_ed_mg_m3_4", substanceRecordset("vla_ed_mg_m3_4").Value
	substance.Add "vla_ec_ppm_4", substanceRecordset("vla_ec_ppm_4").Value
	substance.Add "vla_ec_mg_m3_4", substanceRecordset("vla_ec_mg_m3_4").Value
	substance.Add "notas_vla_4", removeVlbFromNotes(substanceRecordset("notas_vla_4").Value)

	substance.Add "estado_5", substanceRecordset("estado_5").Value
	substance.Add "vla_ed_ppm_5", substanceRecordset("vla_ed_ppm_5").Value
	substance.Add "vla_ed_mg_m3_5", substanceRecordset("vla_ed_mg_m3_5").Value
	substance.Add "vla_ec_ppm_5", substanceRecordset("vla_ec_ppm_5").Value
	substance.Add "vla_ec_mg_m3_5", substanceRecordset("vla_ec_mg_m3_5").Value
	substance.Add "notas_vla_5", removeVlbFromNotes(substanceRecordset("notas_vla_5").Value)

	substance.Add "estado_6", substanceRecordset("estado_6").Value
	substance.Add "vla_ed_ppm_6", substanceRecordset("vla_ed_ppm_6").Value
	substance.Add "vla_ed_mg_m3_6", substanceRecordset("vla_ed_mg_m3_6").Value
	substance.Add "vla_ec_ppm_6", substanceRecordset("vla_ec_ppm_6").Value
	substance.Add "vla_ec_mg_m3_6", substanceRecordset("vla_ec_mg_m3_6").Value
	substance.Add "notas_vla_6", removeVlbFromNotes(substanceRecordset("notas_vla_6").Value)
	
	substance.Add "ib_1", substanceRecordset("ib_1").Value
	substance.Add "vlb_1", substanceRecordset("vlb_1").Value
	substance.Add "momento_1", substanceRecordset("momento_1").Value
	substance.Add "notas_vlb_1", substanceRecordset("notas_vlb_1").Value

	substance.Add "ib_2", substanceRecordset("ib_2").Value
	substance.Add "vlb_2", substanceRecordset("vlb_2").Value
	substance.Add "momento_2", substanceRecordset("momento_2").Value
	substance.Add "notas_vlb_2", substanceRecordset("notas_vlb_2").Value

	substance.Add "ib_3", substanceRecordset("ib_3").Value
	substance.Add "vlb_3", substanceRecordset("vlb_3").Value
	substance.Add "momento_3", substanceRecordset("momento_3").Value
	substance.Add "notas_vlb_3", substanceRecordset("notas_vlb_3").Value

	substance.Add "ib_4", substanceRecordset("ib_4").Value
	substance.Add "vlb_4", substanceRecordset("vlb_4").Value
	substance.Add "momento_4", substanceRecordset("momento_4").Value
	substance.Add "notas_vlb_4", substanceRecordset("notas_vlb_4").Value

	substance.Add "ib_5", substanceRecordset("ib_5").Value
	substance.Add "vlb_5", substanceRecordset("vlb_5").Value
	substance.Add "momento_5", substanceRecordset("momento_5").Value
	substance.Add "notas_vlb_5", substanceRecordset("notas_vlb_5").Value

	substance.Add "ib_6", substanceRecordset("ib_6").Value
	substance.Add "vlb_6", substanceRecordset("vlb_6").Value
	substance.Add "momento_6", substanceRecordset("momento_6").Value
	substance.Add "notas_vlb_6", substanceRecordset("notas_vlb_6").Value

	' Cancer
	substance.Add "notas_cancer_rd", replaceValidated(substanceRecordset("notas_cancer_rd").Value, "v?ase Tabla 3", "")
	substance.Add "grupo_iarc", substanceRecordset("grupo_iarc").Value
	substance.Add "volumen_iarc", substanceRecordset("volumen_iarc").Value
	substance.Add "notas_iarc", substanceRecordset("notas_iarc").Value
	substance.Add "categoria_cancer_otras", substanceRecordset("categoria_cancer_otras").Value
	substance.Add "fuente", substanceRecordset("fuente").Value

	' Disruptor endocrino
	substance.Add "nivel_disruptor", substanceRecordset("nivel_disruptor").Value

	' Neurot?xico
	substance.Add "efecto_neurotoxico", substanceRecordset("efecto_neurotoxico").Value
	substance.Add "nivel_neurotoxico", substanceRecordset("nivel_neurotoxico").Value
	substance.Add "fuente_neurotoxico", substanceRecordset("fuente_neurotoxico").Value

	' TPB
	substance.Add "enlace_tpb", substanceRecordset("enlace_tpb").Value
	substance.Add "anchor_tpb", substanceRecordset("anchor_tpb").Value
	substance.Add "fuente_tpb", substanceRecordset("fuentes_tpb").Value

	' mPmB
	substance.Add "mpmb", substanceRecordset("mpmb").Value

	' Tóxica para el agua
	substance.Add "directiva_aguas", substanceRecordset("directiva_aguas").Value
	substance.Add "clasif_mma", substanceRecordset("clasif_mma").Value
	substance.Add "sustancia_prioritaria", substanceRecordset("sustancia_prioritaria").Value

	' Contaminante del aire
	substance.Add "dano_calidad_aire", substanceRecordset("dano_calidad_aire").Value
	substance.Add "dano_ozono", substanceRecordset("dano_ozono").Value
	substance.Add "dano_cambio_clima", substanceRecordset("dano_cambio_clima").Value

	substance.Add "comentarios_medio_ambiente", substanceRecordset("comentarios_ma").Value

	' Cancer Mama
	substance.Add "cancer_mama_fuente", substanceRecordset("cancer_mama_fuente").Value

	' COP
	substance.Add "cop", substanceRecordset("cop").Value
	substance.Add "enlace_cop", substanceRecordset("enlace_cop").Value

	substance.Add "frasesR", joinFrases("R", substance)

	substance.Add "sinonimos", obtainSynonyms(id_sustancia, connection)
	substance.Add "featuredLists", obtainFeaturedLists(id_sustancia, connection)

	dim substanceGroupsRecordset : set substanceGroupsRecordset = getRecordsetSubstanceGroups(id_sustancia, connection)
	substance.Add "grupos", extractSubstanceGroups(substanceGroupsRecordset)
	set substance = addSubstanceGroupsAssociatedFields(substance, substanceGroupsRecordset)
	substanceGroupsRecordset.close()
	set substanceGroupsRecordset = nothing

	substance.Add "aplicaciones", findSubstanceApplications(id_sustancia, connection)
	substance.Add "compañias", findSubstanceCompanies(id_sustancia, connection)

	substance.Add "pictogramasRd1272", findPictograms(substance.item("simbolos_rd1272"), connection)
	substance.Add "clasificacionesRd1272", findClasificacionesRd1272(substance, connection)
	substance.Add "concentracionEtiquetadoRd1272", obtainConcentracionEtiquetadoRd1272(substance)
	substance.Add _
		"listaNegraClassifications" _
		, getListaNegraClassifications( _
			substance("featuredLists"), substance("frasesR"), substanceRecordset("mpmb") _
		)
	
	set extractSubstance = substance
end function

function extractSubstanceLevelOneFields(substanceId, substanceDic, connection)
	set substance = Server.CreateObject("Scripting.Dictionary")

	substance.Add "nombre", substanceDic("nombre")
	substance.Add "sinonimos", obtainSynonyms(substanceId, connection)
	substance.Add "num_cas", substanceDic("num_cas")
	substance.Add "num_ce_einecs", substanceDic("num_ce_einecs")
	substance.Add "num_ce_elincs", substanceDic("num_ce_elincs")
	substance.Add "num_rd", substanceDic("num_rd")
	substance.Add "nums_icsc", obtainNumsIcsc(substanceDic("num_icsc"))
	substance.Add "pictogramasRd1272", findPictograms(substanceDic("simbolos_rd1272"), connection)
	substance.Add "clasificacionesRd1272", findClasificacionesRd1272(substanceDic, connection)
	substance.Add "notas_rd1272", obtainNotasRd1272(substanceDic("notas_rd1272"), connection)
	substance.Add "concentracionEtiquetadoRd1272", obtainConcentracionEtiquetadoRd1272(substanceDic)

	substanceDic("notas_vla_1") = removeVlbFromNotes(substanceDic("notas_vla_1"))
	substanceDic("notas_vla_2") = removeVlbFromNotes(substanceDic("notas_vla_2"))
	substanceDic("notas_vla_3") = removeVlbFromNotes(substanceDic("notas_vla_3"))
	substanceDic("notas_vla_4") = removeVlbFromNotes(substanceDic("notas_vla_4"))
	substanceDic("notas_vla_5") = removeVlbFromNotes(substanceDic("notas_vla_5"))
	substanceDic("notas_vla_6") = removeVlbFromNotes(substanceDic("notas_vla_6"))
	substance.Add "valoresLimiteAmbiental", obtainValoresLimiteAmbiental(substanceDic, connection)
	substance.Add "valoresLimiteBiologico", obtainValoresLimiteBiologico(substanceDic, connection)
	dim substanceGroupsRecordset : set substanceGroupsRecordset = getRecordsetSubstanceGroups(substanceId, connection)
	substance.Add "grupos", extractSubstanceGroups(substanceGroupsRecordset)
	set substance = addSubstanceGroupsAssociatedFields(substance, substanceGroupsRecordset)
	substanceGroupsRecordset.close()
	set substanceGroupsRecordset = nothing
	substance.Add "aplicaciones", findSubstanceApplications(substanceId, connection)
	substance.Add "featuredLists", obtainFeaturedLists(substanceId, connection)
	substance.Add "frasesRSrz", joinFrases("R", substanceDic)
	substance.Add _
		"listaNegraClassifications" _
		, getListaNegraClassifications( _
			substance("featuredLists"), substance("frasesRSrz"), substanceDic("mpmb") _
		)
	substance.add "explosiva", isSubstanceExplosive(substanceDic("clasificacion_rd1272_1"))
	substance.add "pictogramasRd363", findPictograms(substanceDic("simbolos"), connection)
	substance.add "frasesR", findFrasesR(substance("frasesRSrz"), connection)
	substance.add "frasesRDanesa", findFrasesR( _
		joinFrasesRDanesa( _
			substanceDic("frases_r_danesa") _
		)_
	, connection)
	substance.add "frasesS", findFrasesS(substanceDic("frases_s"), connection)
'	substance.add "concentracionEtiquetadoRd363", obtainConcentracionEtiquetadoRd363(substanceDic, connection)
'	substance.add "notas_rd_363", obtainNotasRd363("notas_rd_363")

	set extractSubstanceLevelOneFields = substance
end function

function extractSubstanceSaludFields(substanceId, substanceDic, connection)
	dim substance : set substance = Server.CreateObject("Scripting.Dictionary")
	dim featuredLists : featuredLists = obtainFeaturedLists(substanceId, connection)
	dim substanceGroupsRecordset : set substanceGroupsRecordset = getRecordsetSubstanceGroups(substanceId, connection)
	set substanceDic = addSubstanceGroupsAssociatedFields(substanceDic, substanceGroupsRecordset)
	substanceGroupsRecordset.close()
	set substanceGroupsRecordset = nothing
	substance.add "grupo_iarc", extractGrupoIarc(substanceDic("grupo_iarc"))
	substance.add "volumen_iarc", substanceDic("volumen_iarc")
	substance.add "notas_iarc", substanceDic("notas_iarc")
	substance.add "nivel_disruptor", obtainDefinitions(substanceDic("nivel_disruptor"), connection)
	substance.add "efecto_neurotoxico", obtainEfectosNeurotoxico(substanceDic("efecto_neurotoxico"), featuredLists, connection)
	substance.add "fuente_neurotoxico", obtainFuentesNeurotoxico(substanceDic("fuente_neurotoxico"), featuredLists, connection)
	dim nivel_neurotoxico_key
	nivel_neurotoxico_key = obtainNivelNeurotoxicoKey(substanceDic("nivel_neurotoxico"))
	substance.add "nivel_neurotoxico", obtainDefinitions(nivel_neurotoxico_key, connection)
	substance.add "nivel_tpr", obtainNivelTpr(substanceDic, connection)

	set extractSubstanceSaludFields = substance
end function

function extractSubstanceMedioAmbienteFields(substanceId, substanceDic, connection)
	dim substance : set substance = Server.CreateObject("Scripting.Dictionary")
	dim substanceGroupsRecordset : set substanceGroupsRecordset = getRecordsetSubstanceGroups(substanceId, connection)
	set substanceDic = addSubstanceGroupsAssociatedFields(substanceDic, substanceGroupsRecordset)
	substanceGroupsRecordset.close()
	set substanceGroupsRecordset = nothing
	substance.add "anchor_tpb", substanceDic("anchor_tpb")
	substance.add "enlace_tpb", substanceDic("enlace_tpb")
	substance.add "fuentes_tpb", obtainDefinitions(substanceDic("fuentes_tpb"), connection)
	substance.add "directiva_aguas", substanceDic("directiva_aguas")
	substance.add "clasif_mma", obtainClasifMma(substanceDic("clasif_mma"), connection)
	substance.add "sustancia_prioritaria", substanceDic("sustancia_prioritaria")

	set extractSubstanceMedioAmbienteFields = substance
end function

function extractCancerOtrasFields(substanceId, substanceDic, connection)
	dim substance : set substance = Server.CreateObject("Scripting.Dictionary")
	dim substanceGroupsRecordset : set substanceGroupsRecordset = getRecordsetSubstanceGroups(substanceId, connection)
	set substanceDic = addSubstanceGroupsAssociatedFields(substanceDic, substanceGroupsRecordset)
	substanceGroupsRecordset.close()
	substance.add "categorias_cancer_otras", obtainCategoriasCancerOtras( _
		substanceDic("categoria_cancer_otras") _
		, substanceDic("fuente") _
		, connection )

	set extractCancerOtrasFields = substance
end function

function extractEnfermedadesFields(substanceDics)
	dim substance : set substance = Server.CreateObject("Scripting.Dictionary")
	dim enfermedades : enfermedades = obtainEnfermedades(substanceDics)
	substance.add "enfermedades_profesionales", enfermedades

	set extractEnfermedadesFields = substance
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

function composeSubstanceQuery(id_sustancia)
	sql = "SELECT *,dn_risc_sustancias_ambiente.comentarios as comentarios_ma, dn_risc_sustancias.comentarios as comentarios_sustancia "
	sql = sql & " FROM dn_risc_sustancias  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_vl ON dn_risc_sustancias.id = dn_risc_sustancias_vl.id_sustancia  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_iarc ON dn_risc_sustancias.id = dn_risc_sustancias_iarc.id_sustancia  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_cancer_otras ON dn_risc_sustancias.id = dn_risc_sustancias_cancer_otras.id_sustancia  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_neuro_disruptor ON dn_risc_sustancias.id = dn_risc_sustancias_neuro_disruptor.id_sustancia  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_ambiente ON dn_risc_sustancias.id = dn_risc_sustancias_ambiente.id_sustancia  "
	sql = sql & " FULL OUTER JOIN dn_risc_sustancias_mama_cop ON dn_risc_sustancias.id = dn_risc_sustancias_mama_cop.id_sustancia  "
	sql = sql & " FULL OUTER JOIN spl_risc_sustancias_prohibidas_embarazadas ON dn_risc_sustancias.id = spl_risc_sustancias_prohibidas_embarazadas.id_sustancia  "
	sql = sql & " WHERE dn_risc_sustancias.id="&id_sustancia
	composeSubstanceQuery = sql
end function

function composeSubstanceLevelOneFieldsQuery(id_sustancia)
	sql = _
		"SELECT " &_
			"sus.id as substanceId, sus.nombre, sus.num_cas, sus.num_ce_einecs, sus.num_ce_elincs, sus.num_rd, sus.num_icsc, " &_
			"sus.simbolos_rd1272, sus.notas_rd1272, " &_
			"sus.clasificacion_rd1272_1, sus.clasificacion_rd1272_2, sus.clasificacion_rd1272_3, " &_
			"sus.clasificacion_rd1272_4, sus.clasificacion_rd1272_5, sus.clasificacion_rd1272_6, " &_
			"sus.clasificacion_rd1272_7, sus.clasificacion_rd1272_8, sus.clasificacion_rd1272_9, " &_
			"sus.clasificacion_rd1272_10, sus.clasificacion_rd1272_11, sus.clasificacion_rd1272_12, " &_
			"sus.clasificacion_rd1272_13, sus.clasificacion_rd1272_14, sus.clasificacion_rd1272_15," &_
			"sus.conc_rd1272_1, sus.conc_rd1272_2, sus.conc_rd1272_3, " &_
			"sus.conc_rd1272_4, sus.conc_rd1272_5, sus.conc_rd1272_6, " &_
			"sus.conc_rd1272_7, sus.conc_rd1272_8, sus.conc_rd1272_9, " &_
			"sus.conc_rd1272_10, sus.conc_rd1272_11, sus.conc_rd1272_12, " &_
			"sus.conc_rd1272_13, sus.conc_rd1272_14, sus.conc_rd1272_15, " &_
			"sus.eti_conc_rd1272_1, sus.eti_conc_rd1272_2, sus.eti_conc_rd1272_3, " &_
			"sus.eti_conc_rd1272_4, sus.eti_conc_rd1272_5, sus.eti_conc_rd1272_6, " &_
			"sus.eti_conc_rd1272_7, sus.eti_conc_rd1272_8, sus.eti_conc_rd1272_9, " &_
			"sus.eti_conc_rd1272_10, sus.eti_conc_rd1272_11, sus.eti_conc_rd1272_12, " &_
			"sus.eti_conc_rd1272_13, sus.eti_conc_rd1272_14, sus.eti_conc_rd1272_15, " &_
			"sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, " &_
			"sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, " &_
			"sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, " &_
			"sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, " &_
			"sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15, " &_
			"sus_vl.estado_1, sus_vl.estado_2, sus_vl.estado_3, sus_vl.estado_4, sus_vl.estado_5, sus_vl.estado_6, " &_
			"sus_vl.vla_ed_ppm_1, sus_vl.vla_ed_ppm_2, sus_vl.vla_ed_ppm_3, " &_
			"sus_vl.vla_ed_ppm_4, sus_vl.vla_ed_ppm_5, sus_vl.vla_ed_ppm_6, " &_
			"sus_vl.vla_ed_mg_m3_1, sus_vl.vla_ed_mg_m3_2, sus_vl.vla_ed_mg_m3_3, " &_
			"sus_vl.vla_ed_mg_m3_4, sus_vl.vla_ed_mg_m3_5, sus_vl.vla_ed_mg_m3_6, " &_
			"sus_vl.vla_ec_ppm_1, sus_vl.vla_ec_ppm_2, sus_vl.vla_ec_ppm_3, " &_
			"sus_vl.vla_ec_ppm_4, sus_vl.vla_ec_ppm_5, sus_vl.vla_ec_ppm_6, " &_
			"sus_vl.vla_ec_mg_m3_1, sus_vl.vla_ec_mg_m3_2, sus_vl.vla_ec_mg_m3_3, " &_
			"sus_vl.vla_ec_mg_m3_4, sus_vl.vla_ec_mg_m3_5, sus_vl.vla_ec_mg_m3_6, " &_
			"sus_vl.notas_vla_1, sus_vl.notas_vla_2, sus_vl.notas_vla_3, " &_
			"sus_vl.notas_vla_4, sus_vl.notas_vla_5, sus_vl.notas_vla_6, " &_
			"sus_vl.ib_1, sus_vl.vlb_1, sus_vl.momento_1, sus_vl.notas_vlb_1, " &_
			"sus_vl.ib_2, sus_vl.vlb_2, sus_vl.momento_2, sus_vl.notas_vlb_2, " &_
			"sus_vl.ib_3, sus_vl.vlb_3, sus_vl.momento_3, sus_vl.notas_vlb_3, " &_
			"sus_vl.ib_4, sus_vl.vlb_4, sus_vl.momento_4, sus_vl.notas_vlb_4, " &_
			"sus_vl.ib_5, sus_vl.vlb_5, sus_vl.momento_5, sus_vl.notas_vlb_5, " &_
			"sus_vl.ib_6, sus_vl.vlb_6, sus_vl.momento_6, sus_vl.notas_vlb_6, " &_
			"sus_amb.mpmb" &_
			", sus.simbolos, sus.frases_r_danesa, sus.frases_s " &_
		"FROM " &_
			"dn_risc_sustancias as sus " &_
		"LEFT JOIN dn_risc_sustancias_vl as sus_vl " &_
			"ON sus.id = sus_vl.id_sustancia " &_
		"LEFT JOIN dn_risc_sustancias_ambiente as sus_amb " &_
			"ON sus.id = sus_amb.id_sustancia " &_
		"WHERE sus.id = " & id_sustancia

	composeSubstanceLevelOneFieldsQuery = sql
end function

function composeSaludQuery(id_sustancia)
	dim sql
	sql = _
		"SELECT " &_
			"sus.id, iarc.grupo_iarc, iarc.notas_iarc, iarc.volumen_iarc, " &_
			"neurodis.nivel_disruptor, neurodis.efecto_neurotoxico, neurodis.fuente_neurotoxico, neurodis.nivel_neurotoxico, " &_
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
		"WHERE " &_
			"sus.id = " & id_sustancia

	composeSaludQuery = sql
end function

function composeMedioAmbienteQuery(substanceId)
	dim sql
	sql = _
		"SELECT " &_
			"anchor_tpb, enlace_tpb, fuentes_tpb, " &_
			"directiva_aguas, clasif_mma, sustancia_prioritaria " &_
		"FROM " &_
			"dn_risc_sustancias_ambiente " &_
		"WHERE " &_
			"id_sustancia = " & substanceId

	composeMedioAmbienteQuery = sql
end function

function composeCancerOtrasQuery(substanceId)
	dim sql
	sql = _
		"SELECT " &_
			"categoria_cancer_otras, fuente " &_
		"FROM " &_
			"dn_risc_sustancias_cancer_otras "	 &_
		"WHERE " &_
			"id_sustancia = " & substanceId

	composeCancerOtrasQuery = sql
end function

function composeEnfermedadesQuery(substanceId)
	dim sql
	sql = _
		"SELECT " &_
		"DISTINCT enf.id, enf.listado, enf.nombre, enf.sintomas, enf.actividades " &_
		"FROM " &_
			"dn_risc_enfermedades AS enf " &_
			"LEFT JOIN dn_risc_grupos_por_enfermedades AS gpe " &_
				"ON enf.id = gpe.id_enfermedad " &_
			"LEFT JOIN dn_risc_sustancias_por_grupos AS spg " &_
				"ON gpe.id_grupo = spg.id_grupo " &_
			"LEFT JOIN dn_risc_sustancias_por_enfermedades AS spe " &_
				"ON spe.id_enfermedad = enf.id " &_
			"WHERE " &_
				"spg.id_sustancia = " & substanceId & " " &_ 
				"OR spe.id_sustancia = " & substanceId & " " &_
			"ORDER BY " &_
				"enf.listado, enf.nombre"
	composeEnfermedadesQuery = sql
end function

function removeVlbFromNotes(notes)
	res = notes
	if (not isnull(notes)) then
		res = replaceValidated(notes, "VLB", "")
	end if

	removeVlbFromNotes = res
end function

function replaceValidated(sourceString, targetString, replaceString)
	output = ""
	if (not isNull(sourceString)) then output = replace(sourceString, targetString, replaceString)

	replaceValidated = output
end function

function joinFrases(tipo, substance)

	' Cada llamada va concatenando a las frases acumuladas anteriormente
	frases = ""
	frases = extractFrase(substance("clasificacion_1"), frases, tipo)
	frases = extractFrase(substance("clasificacion_2"), frases, tipo)
	frases = extractFrase(substance("clasificacion_3"), frases, tipo)
	frases = extractFrase(substance("clasificacion_4"), frases, tipo)
	frases = extractFrase(substance("clasificacion_5"), frases, tipo)
	frases = extractFrase(substance("clasificacion_6"), frases, tipo)
	frases = extractFrase(substance("clasificacion_7"), frases, tipo)
	frases = extractFrase(substance("clasificacion_8"), frases, tipo)
	frases = extractFrase(substance("clasificacion_9"), frases, tipo)
	frases = extractFrase(substance("clasificacion_10"), frases, tipo)
	frases = extractFrase(substance("clasificacion_11"), frases, tipo)
	frases = extractFrase(substance("clasificacion_12"), frases, tipo)
	frases = extractFrase(substance("clasificacion_13"), frases, tipo)
	frases = extractFrase(substance("clasificacion_14"), frases, tipo)
	frases = extractFrase(substance("clasificacion_15"), frases, tipo)

	joinFrases=frases
end function

function joinFrasesRDanesa(byval frases_r)
	' Las frases R danesas vienen separadas por espacios, y para cada una si tiene simbolo, separado por punto y coma

	frases = ""
	array_1 = split (frases_r, " ")
	for i=0 to ubound(array_1)
		'response.write "<br />"&array_1(i)
		' Para cada frase sustituimos punto y coma por espacio para usar el mismo formato que RD y poder extraer la frase
		array_1(i) = replace(array_1(i), ";", " ")
		'response.write "<br />"&array_1(i)
		frases = extractFrase(array_1(i), frases, "R")
		'response.write "<br />"&frases
	next

	' Devolvemos las frases R danesas
	joinFrasesRDanesa = frases
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

function extractFrase(c,f, tipo)
	' Saca las frases R, eliminando el símbolo

	' Arreglamos la frase en caso de que tenga "-" y "/"
	c=formatFrases(c, tipo)

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

function getListaNegraClassifications(featuredLists, frasesR, mpmb)
	dim result : result = Array()

	if anyElementInArray(Array _
		( "cancer_danesa" _
		, "cancer_iarc_excepto_grupo_3" _
		, "cancer_otras_excepto_grupo_4" _
		, "cancer_mama" _
		), featuredLists) then arrayPush result, "cancerígena"
	if inArray("cop", featuredLists) then
		arrayPush result, "cop"
	end if
	if anyElementInArray(Array _
		("mutageno_rd", "mutageno_danesa"), featuredLists) then	arrayPush result, "mutágena"
	if inArray("de", featuredLists) then
		arrayPush result, "disruptora endocrina"
	end if
	if anyElementInArray(Array _
		( "neurotoxico" _
		, "neurotoxico_rd" _
		, "neurotoxico_danesa" _
		, "neurotoxico_nivel" _
		), featuredLists) then arrayPush result, "neurotóxica"
	if anyElementInArray(Array _
		( "sensibilizante" _
		, "sensibilizante_danesa" _
		, "sensibilizante_reach" _
		), featuredLists) then arrayPush result, "sensibilizante"
	if anyElementInArray(Array _
		("tpr", "tpr_danesa" _ 
		), featuredLists) then arrayPush result, "tóxica para la reproducción"
	if stringContains(frasesR, "R33") then
		arrayPush result, "bioacumulativa"
	end if
	if stringContains(frasesR, "R58") then
		arrayPush result, "puede provocar a largo plazo efectos negativos en el medio ambiente"
	end if
	if inArray("tpb", featuredLists) then
		arrayPush result, "tóxica, persistente y bioacumulativa"
	end if
	if mpmb then
		arrayPush result, "muy persistente y muy bioacumulativa"
	end if
	if stringContains(frasesR, "R53") _
		or stringContains(frasesR, "R50-53") _
		or stringContains(frasesR, "R51-53") _
		or stringContains(frasesR, "R52-53") _
		then arrayPush result, "puede provocar a largo plazo efectos negativos en el medio ambiente acuático"

	getListaNegraClassifications = result
end function

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

function isSubstanceExplosive(clasificacion_rd1272_1)
	isSubstanceExplosive = false
	if clasificacion_rd1272_1 = "Expl., ****;" then isSubstanceExplosive = true
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

function obtainEfectosNeurotoxico(byVal efectosSrz, featuredLists, connection)
	obtainEfectosNeurotoxico = efectosSrz
	if isNull(efectosSrz) then _
		exit function
	obtainEfectosNeurotoxico = obtainDefinitions( _ 
		replace(efectosSrz, "/", ",") _
		, connection)
	dim efectos : efectos = split(efectosSrz, "/")
	if not( _
		inArray("neurotoxico_rd", featuredLists) _
		or inArray("neurotoxico_danesa", featuredLists) _
		) then exit function
	if not inArray("SNC", efectos) then _
		arrayPush efectos, "SNC"

	obtainEfectosNeurotoxico = obtainDefinitions( _
		join(efectos, ",") _
		, connection )
end function

function obtainFuentesNeurotoxico(fuentesSrz, featuredLists, connection)
	obtainFuentesNeurotoxico = obtainDefinitions(fuentesSrz, connection)
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
		, connection )
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
	, connection)
end function

function obtainNivelNeurotoxicoKey(nivel)
	obtainNivelNeurotoxicoKey = nivel
	if isNull(nivel) or nivel = "" then _
		exit function
	
	obtainNivelNeurotoxicoKey = "Nivel " &  nivel
end function

function obtainCategoriasCancerOtras(categoriasSrz, fuentesSrz, connection)
	obtainCategoriasCancerOtras = Array()
	if isNull(categoriasSrz) _
		or categoriasSrz = "" _
		or isNull(fuentesSrz) _
		or fuentesSrz = "" then _
		exit function
	dim element
	dim categorias : categorias = split(categoriasSrz, ", ")
	dim fuentes : fuentes = split(fuentesSrz, ", ")
	dim i
	for i = 0 to Ubound(categorias)
		set element = Server.CreateObject("Scripting.Dictionary")
		element.add "categoria", obtainDefinitions(categorias(i), connection)
		element.add "fuente", fuentes(i)
		obtainCategoriasCancerOtras = arrayPushDictionary(obtainCategoriasCancerOtras, element)
	next
end function

function obtainEnfermedades(substanceDics)
	dim result : result = Array()
	dim key, enfermedad
	for each key in substanceDics
		set enfermedad = Server.CreateObject("Scripting.Dictionary")
		enfermedad.add "id", key("id")
		enfermedad.add "listado", formatEnfermedades(key("listado"))
		enfermedad.add "nombre", formatEnfermedades(key("nombre"))
		enfermedad.add "sintomas", crLfToBr(key("sintomas"))
		enfermedad.add "actividades", crLfToBr(key("actividades"))
		result = arrayPushDictionary(result, enfermedad)
	next
	
	obtainEnfermedades = result
end function

function formatEnfermedades(value)
	formatEnfermedades = ""
	if isEmpty(value) then _
		exit function
	formatEnfermedades = capitalize( _
		trim( _
			replace( _
				replace(value, "Enfermedades", "") _
			, "profesionales", "") _
		) _
	)
end function

function capitalize(str)
	capitalize = UCase(Left(str, 1)) & Mid(str, 2)
end function

function crLfToBr(value)
	crLfToBr = replace(value, vbCrlf, "<br>")
end function

function obtainClasifMma(byVal clasif, connection)
	dim key : key = ""
	obtainClasifMma = ""
	if not isNumeric(clasif) then _
		exit function
	if (0 < clasif < 4 ) then _
		key = clasif + "."

	obtainClasifMma = obtainDefinitions(key, connection)
end function
%>