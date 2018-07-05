<%
function addSubstanceGroupsAssociatedFields(substance, substanceGroupsRecordset)
	dim substanceTables

	set substanceTables = collectSubstanceTables()

	do while not substanceGroupsRecordset.eof
		for each list in substanceTables.keys
			set substance = evaluaCamposListaAsociada(substance, substanceGroupsRecordset, list, substanceTables.Item(list))
		next
		substanceGroupsRecordset.movenext
	loop

	set addSubstanceGroupsAssociatedFields = substance
end function

function extractSubstanceGroups(substanceGroupsRecordset)
	dim result : result = Array()
	dim group

	if substanceGroupsRecordset.Eof then
		exctratSubstanceGroups = result
		exit function
	end if
	do while not substanceGroupsRecordset.Eof
		set group = Server.CreateObject("Scripting.Dictionary")
		group.add "item_id", substanceGroupsRecordset("item_id").value
		group.add "name", substanceGroupsRecordset("name").value
		group.add "description", substanceGroupsRecordset("description").value
		result = arrayPushDictionary(result, group)
		set group = nothing
		substanceGroupsRecordset.MoveNext
	loop
	if substanceGroupsRecordset.Eof then substanceGroupsRecordset.MoveFirst

	extractSubstanceGroups = result
end function

' PRIVATE
function getRecordsetSubstanceGroupsInternational(id_sustancia, lang, connection)
	dim sqlQuery
	if lang = "en" then
		sqlQuery = "SELECT gr.*, gr.id AS item_id, gr.nombre_ing as name, gr.descripcion_ing as description FROM dn_risc_grupos gr, dn_risc_sustancias_por_grupos sg WHERE sg.id_grupo=gr.id AND sg.id_sustancia=" & id_sustancia & " ORDER BY nombre"
		set getRecordsetSubstanceGroupsInternational = connection.execute(sqlQuery)
		exit function
	end if

	sqlQuery = "SELECT gr.*, gr.id AS item_id, gr.nombre as name, gr.descripcion as description FROM dn_risc_grupos gr, dn_risc_sustancias_por_grupos sg WHERE sg.id_grupo=gr.id AND sg.id_sustancia=" & id_sustancia & " ORDER BY nombre"
	set getRecordsetSubstanceGroupsInternational = connection.execute(sqlQuery)
end function

function evaluaCamposListaAsociada(substance, substanceGroupsRecordset, listName, groupKeys)
	dim currentFieldName, currentSubstanceGroupValue, currentSubstanceValue
	dim fieldName
	fieldName = "asoc_" & listName
	if not FieldExists(substanceGroupsRecordset, fieldName) then
		set evaluaCamposListaAsociada = substance
		exit function
	end if
	for i = 0 to UBound(groupKeys)
		currentGroupKey = groupKeys(i)
		currentSubstanceValue = getNonNullValue(substance(currentGroupKey))
		currentFieldName = fieldName & "_" & currentGroupKey
		if FieldExists(substanceGroupsRecordset, currentFieldName) then
			currentSubstanceGroupValue = getNonNullValue(substanceGroupsRecordset(currentFieldName))
			substance(currentGroupKey) = appendNotPresentValue(currentSubstanceValue, currentSubstanceGroupValue)
		end if
	next

	set evaluaCamposListaAsociada = substance
end function

function getNonNullValue(value)
	getNonNullValue = ""
	if isNull(value) then _
		exit function

	getNonNullValue = value
end function

function appendNotPresentValue(byVal valuesSrz1, byVal valuesSrz2)
	appendNotPresentValue = valuesSrz1
	if is_empty(valuesSrz2) then _
		exit function
	if is_empty(valuesSrz1) then
		appendNotPresentValue = valuesSrz2
		exit function
	end if
	dim i
	dim valuesSrz1LCased : valuesSrz1LCased = lCase(valuesSrz1)
	dim values1 : values1 = split(valuesSrz1, ", ")
	dim values1LCased : values1LCased = split(valuesSrz1LCased, ", ")
	dim values2 : values2 = split(valuesSrz2, ", ")
	for i = 0 to UBound(values2)
		if not inArray( _
			lcase(values2(i)) _
			, values1LCased) then _
				arrayPush values1, values2(i)
	next

	appendNotPresentValue = join(values1, ", ")
end function

function collectSubstanceTables()
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

Function FieldExists(ByVal rs, ByVal fieldName)

    On Error Resume Next
    FieldExists = rs.Fields(fieldName).name <> ""
    If Err <> 0 Then FieldExists = False
    Err.Clear

End Function
%>
