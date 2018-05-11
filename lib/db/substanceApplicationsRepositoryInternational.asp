<%
function findSubstanceApplicationsInternational(substance_id, lang, connection)
	dim sqlQuery, substanceApplicationsRecordset

	sqlQuery = composeSubtanceApplicationsQueryInternational(substance_id, lang)
	set substanceApplicationsRecordset = connection.execute(sqlQuery)
	findSubstanceApplicationsInternational = extractSubstanceApplications(substanceApplicationsRecordset)

end function

' PRIVATE
function composeSubtanceApplicationsQueryInternational(substance_id, lang)
	dim result
	if lang = "en" then
		result = "SELECT " &_
			"DISTINCT u.id AS item_id, u.nombre_ing AS name, u.descripcion_ing AS description " &_
			"FROM dn_risc_usos AS u " &_
			"LEFT OUTER JOIN dn_risc_grupos_por_usos AS gpu " &_
				"ON u.id = gpu.id_uso " &_
			"LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg " &_
				"ON gpu.id_grupo = spg.id_grupo " &_
			"LEFT OUTER JOIN dn_risc_sustancias_por_usos AS spu " &_
				"ON spu.id_uso = u.id " &_
			"WHERE spg.id_sustancia = " & substance_id & " OR spu.id_sustancia = " & substance_id & " " &_
			"ORDER BY u.nombre_ing"

			composeSubtanceApplicationsQueryInternational = result
			exit function
	end if

	result = "SELECT " &_
		"DISTINCT u.id AS item_id, u.nombre AS name, u.descripcion AS description " &_
		"FROM dn_risc_usos AS u " &_
		"LEFT OUTER JOIN dn_risc_grupos_por_usos AS gpu " &_
			"ON u.id = gpu.id_uso " &_
		"LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg " &_
			"ON gpu.id_grupo = spg.id_grupo " &_
		"LEFT OUTER JOIN dn_risc_sustancias_por_usos AS spu " &_
			"ON spu.id_uso = u.id " &_
		"WHERE spg.id_sustancia = " & substance_id & " OR spu.id_sustancia = " & substance_id & " " &_
		"ORDER BY u.nombre"

	composeSubtanceApplicationsQueryInternational = result
end function

function extractSubstanceApplications(substanceApplicationsRecordset)
	dim result : result = Array()
	dim substanceApplication

	if substanceApplicationsRecordset.Eof then
		extractSubstanceApplications = result
		exit function
	end if
	do while not substanceApplicationsRecordset.Eof
		set substanceApplication = Server.CreateObject("Scripting.Dictionary")
		substanceApplication.add "item_id", substanceApplicationsRecordset("item_id").value
		substanceApplication.add "name", substanceApplicationsRecordset("name").value
		substanceApplication.add "description", substanceApplicationsRecordset("description").value
		result = arrayPushDictionary(result, substanceApplication)
		set substanceApplication = nothing
		substanceApplicationsRecordset.MoveNext
	loop
	if substanceApplicationsRecordset.Eof then substanceApplicationsRecordset.MoveFirst

	extractSubstanceApplications = result
end function

%>
