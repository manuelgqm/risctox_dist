<%
function findSubstanceApplications(id_sustancia, connection)
	dim sqlQuery, substanceApplicationsRecordset

	sqlQuery = composeSubtanceApplicationsQuery(id_sustancia)
	set substanceApplicationsRecordset = connection.execute(sqlQuery)
	findSubstanceApplications = extractSubstanceApplications(substanceApplicationsRecordset)

end function

' PRIVATE
function composeSubtanceApplicationsQuery(id_sustancia)
	dim result
	result = "SELECT DISTINCT u.id AS item_id, u.nombre AS name, u.descripcion AS description FROM dn_risc_usos AS u " &_
				"LEFT OUTER JOIN dn_risc_grupos_por_usos AS gpu " &_
					"ON u.id = gpu.id_uso " &_
				"LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg " &_
					"ON gpu.id_grupo = spg.id_grupo " &_
				"LEFT OUTER JOIN dn_risc_sustancias_por_usos AS spu " &_
					"ON spu.id_uso = u.id " &_
				"WHERE spg.id_sustancia = " & id_sustancia & " OR spu.id_sustancia = " & id_sustancia & " " &_
				"ORDER BY u.nombre"

	composeSubtanceApplicationsQuery = result
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
		result = arrayPush(result, substanceApplication)
		set substanceApplication = nothing
		substanceApplicationsRecordset.MoveNext	
	loop
	if substanceApplicationsRecordset.Eof then substanceApplicationsRecordset.MoveFirst

	extractSubstanceApplications = result
end function

%>