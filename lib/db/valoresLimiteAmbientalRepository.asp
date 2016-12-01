<%
function obtainValoresLimiteAmbiental(substance, connection)
	Dim result : result = Array()
	Dim estados : estados = Array("estado_1", "estado_2", "estado_3", "estado_4", "estado_5", "estado_6")
	Dim ed_ppm : ed_ppm = Array("vla_ed_ppm_1", "vla_ed_ppm_2", "vla_ed_ppm_3", "vla_ed_ppm_4","vla_ed_ppm_5", "vla_ed_ppm_6")
	Dim ed_mg_m3 : ed_mg_m3 = Array("vla_ed_mg_m3_1", "vla_ed_mg_m3_2", "vla_ed_mg_m3_3", "vla_ed_mg_m3_4", "vla_ed_mg_m3_5", "vla_ed_mg_m3_6")
	Dim ec_ppm : ec_ppm = Array("vla_ec_ppm_1", "vla_ec_ppm_2", "vla_ec_ppm_3", "vla_ec_ppm_4","vla_ec_ppm_5", "vla_ec_ppm_6")
	Dim ec_mg_m3 : ec_mg_m3 = Array("vla_ec_mg_m3_1", "vla_ec_mg_m3_2", "vla_ec_mg_m3_3", "vla_ec_mg_m3_4", "vla_ec_mg_m3_5", "vla_ec_mg_m3_6")
	Dim notas : notas = Array("notas_vla_1", "notas_vla_2", "notas_vla_3", "notas_vla_4", "notas_vla_5", "notas_vla_6")
	Dim i, valorLimiteAmbiental
	for i = 0 to ubound(estados)
			set valorLimiteAmbiental = extractValorLimiteAmbiental(substance, connection, _
				estados(i), _
				ed_ppm(i), _
				ed_mg_m3(i), _
				ec_ppm(i), _
				ec_mg_m3(i), _
				notas(i) _
			)
		if not isDictionaryEmpty(valorLimiteAmbiental) then
			result = arrayPushDictionary(result, valorLimiteAmbiental)
		end if
	next

	obtainValoresLimiteAmbiental = result
end function

function extractValorLimiteAmbiental(substance, connection, estado, ed_ppm, ed_mg_m3, ec_ppm, ec_mg_m3, notas)
	Dim result : set result = Server.CreateObject("Scripting.Dictionary")
	result.add "estado", substance.item(estado)
	result.add "ed_ppm", substance.item(ed_ppm)
	result.add "ed_mg_m3", substance.item(ed_mg_m3)
	result.add "ec_ppm", substance.item(ec_ppm)
	result.add "ec_mg_m3", substance.item(ec_mg_m3)
	result.add "notas", obtainNotasValoresLimiteAmbiental(substance.item(notas), connection)

	set extractValorLimiteAmbiental = result
end function

function obtainNotasValoresLimiteAmbiental(byVal notasSrz, connection)
	dim result : result = Array()
	Dim notas
	if isNull(notasSrz) then 
		notasSrz = ""
	end if
	notas = split(notasSrz, ",")
	dim i, nota
	dim notasCleared : notasCleared = array()
	for i = 0 to Ubound(notas)
		nota = replace(notas(i), " ", "")
		if len(nota) then arrayPush notasCleared, nota
	next
	obtainNotasValoresLimiteAmbiental = findDefinitions(notasCleared, connection)
end function
%>