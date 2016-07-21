<%
function findSubstanceCompanies(id_sustancia, connection)
	dim sqlQuery, substanceCompaniesRecordset

	sqlQuery = composeSubtanceCompaniesQuery(id_sustancia)
	set substanceCompaniesRecordset = connection.execute(sqlQuery)
	findSubstanceCompanies = extractSubstanceCompanies(substanceCompaniesRecordset)
end function

function composeSubtanceCompaniesQuery(id_sustancia)
	dim result
	result = "SELECT dn_risc_companias.id as item_id, nombre as name " &_
				"FROM dn_risc_sustancias_por_companias INNER JOIN dn_risc_companias " &_
					"ON dn_risc_sustancias_por_companias.id_compania = dn_risc_companias.id " &_
				"WHERE id_sustancia = " & id_sustancia & " ORDER BY nombre"
	composeSubtanceCompaniesQuery = result
end function

function extractSubstanceCompanies(substanceCompaniesRecordset)
	dim result : result = Array()
	
	if substanceCompaniesRecordset.Eof then
		extractSubstanceCompanies = result
		exit function
	end if
	do while not substanceCompaniesRecordset.Eof
		set substanceCompany = Server.CreateObject("Scripting.Dictionary")
		substanceCompany.add "item_id", substanceCompaniesRecordset("item_id").value
		substanceCompany.add "name", substanceCompaniesRecordset("name").value
		result = arrayPushDictionary(result, substanceCompany)
		set substanceCompany = nothing
		substanceCompaniesRecordset.MoveNext	
	loop
	if substanceCompaniesRecordset.Eof then substanceCompaniesRecordset.MoveFirst

	extractSubstanceCompanies = result
end function
%>