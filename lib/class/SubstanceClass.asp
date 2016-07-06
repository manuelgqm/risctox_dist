<!--#include file="../db/synonymsRepository.asp"-->
<!--#include file="../db/substanceListsRepository.asp"-->
<%
Class SubstanceClass
	Private mNombre
	Private mCollection

	Public property Get nombre()
		nombre = mNombre
	End property
	Public property Let nombre(pData)
		mNombre = pData
	End property

	Public property Get collection()
		set collection = mCollection
	End property
	Public property Let collection(pData)
		set mCollection = pData
	End property

	Public function find(id_sustancia, connection)
		sql = composeSubstanceQuery( id_sustancia )

		set substanceRecordset = connection.execute(sql)
		set substance = extractSubstance(substanceRecordset)
		substanceRecordset.close()
		set substanceRecordset=nothing

		set substance = addSubstanceGroupsAssociatedFields(substance, id_sustancia, connection)

		substance.Add "sinonimos", obtainSynonyms(id_sustancia, connection)
		substance.Add "featuredLists", obtainFeaturedLists(id_sustancia, connection)

		Me.nombre = substance.Item("nombre")
		Me.collection = substance
		find = id_sustancia
	end function

	Private function composeSubstanceQuery(id_sustancia)
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

	Private function extractSubstance(substanceRecordset)
		set substance = Server.CreateObject("Scripting.Dictionary")

		' dn_risc_sustancias
		substance.Add "nombre", substanceRecordset("nombre").Value
		substance.Add "nombre_ing", elimina_repes(substanceRecordset("nombre_ing").Value, "@")

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
		if varType(substanceRecordset("notas_rd1272").Value) = vbNull then
			notas_rd1272 = ""
		else
			notas_rd1272 = substanceRecordset("notas_rd1272").Value
		end if
		substance.Add "notas_rd1272", replaceValidated(substanceRecordset("notas_rd1272"), "@", "@ ")
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

		set extractSubstance = substance
	end function

	Private function addSubstanceGroupsAssociatedFields(substance, id_sustancia, connection)
		dim substanceTables, substanceGroupsRecordset

		set substanceTables = collectSubstanceTables()
		set substanceGroupsRecordset = requestSubstanceGroups(id_sustancia, connection)

		do while not substanceGroupsRecordset.eof
			for each list in substanceTables.keys
				set substance = evaluaCamposListaAsociada(substance, substanceGroupsRecordset, list, substanceTables.Item(list))
			next
			substanceGroupsRecordset.movenext
		loop

		substanceGroupsRecordset.close()
		set substanceGroupsRecordset = nothing

		set addSubstanceGroupsAssociatedFields = substance
	end function

	Private function requestSubstanceGroups(id_sustancia, connection)
		dim sqlQuery

		sqlQuery = "SELECT gr.* FROM dn_risc_grupos gr, dn_risc_sustancias_por_grupos sg WHERE sg.id_grupo=gr.id AND sg.id_sustancia=" & id_sustancia

		set requestSubstanceGroups = connection.execute(sqlQuery)
	end function

	Private function removeVlbFromNotes(notes)
		res = notes
		if (not isnull(notes)) then
			res = replaceValidated(notes, "VLB", "")
		end if

		removeVlbFromNotes = res
	end function

	Private function replaceValidated(sourceString, targetString, replaceString)
		output = ""
		if (not isNull(sourceString)) then output = replace(sourceString, targetString, replaceString)

		replaceValidated = output
	end function

	Private function collectSubstanceTables()
		set lists = Server.CreateObject("Scripting.Dictionary")
		lists.Add "cancer_rd", Array("notas_cancer_rd")
		lists.Add "cancer_iarc", Array("grupo_iarc","volumen_iarc")
		lists.Add "cancer_otras", Array("categoria_cancer_otras","fuente")
		lists.Add "cancer_mama", Array("cancer_mama_fuente")
		lists.Add "neuro_oto", Array("efecto_neurotoxico","nivel_neurotoxico","fuente_neurotoxico")
		lists.Add "disruptores", Array("nivel_disruptor")
		lists.Add "tpb", Array("enlace_tpb","anchor_tpb","fuentes_tpb")
		lists.Add "directiva_aguas", Array("clasif_mma")
		lists.Add "vla", Array("estado_1","ed_ppm_1", "ed_mg_m3_1", "ec_ppm_1", "ec_mg_m3_1", "notas_vla_1",	"estado_2", "ed_ppm_2", "ed_mg_m3_2", "ec_ppm_2", "ec_mg_m3_2", "notas_vla_2", "estado_3", "ed_ppm_3", "ed_mg_m3_3", "ec_ppm_3", "ec_mg_m3_3", "notas_vla_3", "estado_4", "ed_ppm_4", "ed_mg_m3_4", "ec_ppm_4", "ec_mg_m3_4", "notas_vla_4", "estado_5", "ed_ppm_5", "ed_mg_m3_5", "ec_ppm_5", "ec_mg_m3_5", "notas_vla_5", "estado_6", "ed_ppm_6", "ed_mg_m3_6", "ec_ppm_6", "ec_mg_m3_6", "notas_vla_6")
		lists.Add "vlb", Array("ib_1", "vlb_1", "momento_1", "notas_vlb_1", "ib_2", "vlb_2", "momento_2", "notas_vlb_2", "ib_3", "vlb_3", "momento_3", "notas_vlb_3", "ib_4", "vlb_4", "momento_4", "notas_vlb_4", "ib_5", "vlb_5", "momento_5", "notas_vlb_5", "ib_6", "vlb_6", "momento_6", "notas_vlb_6")
		lists.Add "cop", Array("enlace_cop")
		lists.Add "mpmb", Array("")
		lists.Add "eper", Array("")
		lists.Add "eper_agua", Array("")
		lists.Add "eper_aire", Array("")
		lists.Add "eper_suelo", Array("")
		lists.Add "prohibidas", Array("comentario_prohibida")
		lists.Add "restringidas", Array("comentario_restringida")
		lists.Add "prohibidas_embarazadas", Array("comentario_prohibida")
		lists.Add "prohibidas_lactantes", Array("comentario_prohibida")
		lists.Add "candidatas_reach", Array("")
		lists.Add "autorizacion_reach", Array("")
		lists.Add "biocidas_autorizadas", Array("fuente", "pureza_minima", "condiciones", "usos")
		lists.Add "biocidas_prohibidas", Array("fuente", "fecha_limite", "usos")
		lists.Add "pesticidas_autorizadas", Array("fuente", "plazo_renovacion", "pureza_minima", "usos")
		lists.Add "pesticidas_prohibidas", Array("fuente", "exenciones")
		lists.Add "alergeno", Array("")
		lists.Add "calidad_aire", Array("")
		lists.Add  "corap", Array("")
		set collectSubstanceTables = lists
	end function

	Private function joinFrases(tipo, substance)

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

	Private function joinFrasesRDanesa(byval frases_r)
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

	Private function formatFrases(byval c, tipo)
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

	Private function extractFrase(c,f, tipo)
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

End Class
%>