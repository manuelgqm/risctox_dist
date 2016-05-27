<%
function findSubstance( id_sustancia, connection )
	sql = composeSubstanceQuery( id_sustancia )

	set substanceRecordset = connection.execute(sql)
	set substance = extractSubstance(substanceRecordset)
	substanceRecordset.close()
	set substanceRecordset=nothing
	
	set findSubstance = substance
end function

' PRIVATE
function extractSubstance(substanceRecordset)
	set substance = Server.CreateObject("Scripting.Dictionary")
	
	' dn_risc_sustancias
	substance.Add "nombre", substanceRecordset("nombre").Value
	substance.Add "nombre_ing", elimina_repes(substanceRecordset("nombre_ing").Value, "@")
	
	substance.Add "num_rd", substanceRecordset("num_rd").Value
	substance.Add "num_ce_einecs", substanceRecordset("num_ce_einecs").Value
	substance.Add "num_ce_elincs", substanceRecordset("num_ce_elincs").Value
	substance.Add "num_cas", substanceRecordset("num_cas").Value
	substance.Add "cas_alternativos", substanceRecordset("cas_alternativos").Value
	
	' substance.Add "num_onu", substanceRecordset("num_onu").Value 'Parece que no se usa
	
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
	substance.Add "notas_cancer_rd", replaceValidated(substanceRecordset("notas_cancer_rd").Value, "véase Tabla 3", "")
	substance.Add "grupo_iarc", substanceRecordset("grupo_iarc").Value
	substance.Add "volumen_iarc", substanceRecordset("volumen_iarc").Value
	substance.Add "notas_iarc", substanceRecordset("notas_iarc").Value
	substance.Add "categoria_cancer_otras", substanceRecordset("categoria_cancer_otras").Value
	substance.Add "fuente", substanceRecordset("fuente").Value

	' Disruptor endocrino
	substance.Add "nivel_disruptor", substanceRecordset("nivel_disruptor").Value
	'substance.Add "vector_disruptores", split(substance.Item("nivel_disruptor"),",")


	' Neurotóxico
	substance.Add "efecto_neurotoxico", substanceRecordset("efecto_neurotoxico").Value
	substance.Add "nivel_neurotoxico", substanceRecordset("nivel_neurotoxico").Value
	substance.Add "fuente_neurotoxico", substanceRecordset("fuente_neurotoxico").Value

	' TPB
	substance.Add "enlace_tpb", substanceRecordset("enlace_tpb").Value
	substance.Add "anchor_tpb", substanceRecordset("anchor_tpb").Value
	substance.Add "fuente_tpb", substanceRecordset("fuentes_tpb").Value

	' SPL (16/06/2014)
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
	'substance.Add "toxicidad_suelo", substanceRecordset("toxicidad_suelo").Value

	' Sustancia prohibida
	'substance.Add "sustancia_prohibida", substanceRecordset("sustancia_prohibida").Value
	'substance.Add "sustancia_restringida", substanceRecordset("sustancia_restringida").Value
	'substance.Add "comentarios", trim(substanceRecordset("comentarios_sustancia").Value)

	' Cancer Mama
	'substance.Add "cancer_mama", substanceRecordset("cancer_mama").Value
	substance.Add "cancer_mama_fuente", substanceRecordset("cancer_mama_fuente").Value

	' COP
	substance.Add "cop", substanceRecordset("cop").Value
	substance.Add "enlace_cop", substanceRecordset("enlace_cop").Value

	set extractSubstance = substance
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
%>
