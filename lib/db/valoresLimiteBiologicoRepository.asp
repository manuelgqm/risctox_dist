<%
function obtainValoresLimiteBiologico(substance, connection)
	Dim result : result = Array()
	Dim indicadores : indicadores = Array("ib_1", "ib_2", "ib_3", "ib_4", "ib_5", "ib_6")
	Dim valores : valores = Array("vlb_1", "vlb_2", "vlb_3", "vlb_4","vlb_5", "vlb_6")
	Dim momentos : momentos = Array("momento_1", "momento_2", "momento_3", "momento_4", "momento_5", "momento_6")
	Dim notas : notas = Array("notas_vlb_1", "notas_vlb_2", "notas_vlb_3", "notas_vlb_4","notas_vlb_5", "notas_vlb_6")
	Dim i, valorLimiteBiologico
	for i = 0 to ubound(indicadores)
		set valorLimiteBiologico = extractValorLimiteBiologico _
			( substance _
			, connection _
			, indicadores(i) _
			, valores(i) _
			, momentos(i) _
			, notas(i) _
			)
		if not isDictionaryEmpty(valorLimiteBiologico) then
			result = arrayPushDictionary(result, valorLimiteBiologico)
		end if
	next

	obtainValoresLimiteBiologico = result
end function

function extractValorLimiteBiologico(substance, connection, indicador, valor, momento, notas)
	Dim result : set result = Server.CreateObject("Scripting.Dictionary")
	result.add "indicador", substance.item(indicador)
	result.add "valor", substance.item(valor)
	result.add "momento", substance.item(momento)
	result.add "notas", obtainDefinitions(substance.item(notas), connection)

	set extractValorLimiteBiologico = result
end function
%>