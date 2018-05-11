<!--#include file="synonymsRepository.asp"-->
<!--#include file="substanceGroupsRepositoryInternational.asp"-->
<!--#include file="substanceApplicationsRepositoryInternational.asp"-->
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
	
	set extractIdentification = result
end function

function composeIdentificationQuery(substance_id)
	composeIdentificationQuery = _
		"SELECT " &_
			"nombre_ing as nombre, nombre_ing, num_rd, num_ce_einecs, num_ce_elincs, num_cas, " &_
			"cas_alternativos, num_icsc " &_
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
%>
