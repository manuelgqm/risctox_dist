<%
function obtainConcentracionEtiquetadoRd1272(substance)
	Dim etiConcList : etiConcList = Array()
	Dim etiConc : set etiConc = Server.CreateObject("Scripting.Dictionary")
	if substance.item("conc_rd1272_1") = "" and substance.item("conc_rd1272_2") = "" and substance.Item("eti_conc_rd1272_1") = "" Then
		etiConc.add "concentracion", ""
		etiConcList = arrayPushDictionary(etiConcList, etiConc)
		Exit function
	end if
	if substance.item("conc_rd1272_1") = "" and substance.item("conc_rd1272_2") <> "" and substance.Item("eti_conc_rd1272_1") <> "" Then
		etiConc.Add "concentracion", "Factor " & substance.Item("eti_conc_rd1272_1")
		etiConcList = arrayPushDictionary(etiConcList, etiConc)
		Exit function
	end if
	etiConcList = extractEtiConcList(substance)

	obtainConcentracionEtiquetadoRd1272 = etiConcList
end function

function extractEtiConcList(substance)
	Dim result : result = Array()
	Dim etiquetas : etiquetas = Array( _
		"eti_conc_rd1272_1", _
		"eti_conc_rd1272_2", _
		"eti_conc_rd1272_3", _
		"eti_conc_rd1272_4", _
		"eti_conc_rd1272_5", _
		"eti_conc_rd1272_6", _
		"eti_conc_rd1272_7", _
		"eti_conc_rd1272_8", _
		"eti_conc_rd1272_9", _
		"eti_conc_rd1272_10", _
		"eti_conc_rd1272_11", _
		"eti_conc_rd1272_12", _
		"eti_conc_rd1272_13", _
		"eti_conc_rd1272_14", _
		"eti_conc_rd1272_15" _
	)
	Dim concentraciones : concentraciones = Array( _
		"conc_rd1272_1", _
		"conc_rd1272_2", _
		"conc_rd1272_3", _
		"conc_rd1272_4", _
		"conc_rd1272_5", _
		"conc_rd1272_6", _
		"conc_rd1272_7", _
		"conc_rd1272_8", _
		"conc_rd1272_9", _
		"conc_rd1272_10", _
		"conc_rd1272_11", _
		"conc_rd1272_12", _
		"conc_rd1272_13", _
		"conc_rd1272_14", _
		"conc_rd1272_15" _
	)
	Dim etiConc
	Dim i
	for i = 0 to Ubound(etiquetas)
		set etiConc = obtainConcentracionEtiquetado(substance.item(etiquetas(i)), substance.item(concentraciones(i)))
		if etiConc.count > 0 then
			result = arrayPushDictionary(result, etiConc)
		end if
	next

	extractEtiConcList = result
end function

function obtainConcentracionEtiquetado(etiquetado, concentracion)
	dim etiConc : set etiConc = Server.CreateObject("Scripting.Dictionary")
	if isNull(concentracion) and isNull(etiquetado) then
		set obtainConcentracionEtiquetado = etiConc
		exit function
	end if
	'encodeHTMLEntities'
	dim concentracionFormated : concentracionFormated = ""
	dim etiquetadoFormated : etiquetadoFormated = ""
	if etiquetado = "*" Then
		etiquetadoFormated = "Esta entrada tiene límites de concentración específicos para la toxicidad aguda conforme al RD 363/1995 que no pueden «hacerse corresponder» con los límites de concentración con arreglo al Reglamento CLP (como referencia, ver etiquetado del apartado de clasificación (RD 363/1995) de la sustancia)."
		etiConc.add "concentracion", concentracionFormated
		etiConc.add "etiquetado", etiquetadoFormated
		set obtainConcentracionEtiquetado = etiConc
		exit function
	end if
	if concentracion <> "" and etiquetado <> "" Then
		concentracionFormated = replace(concentracion, ":", "")
		concentracionFormated = encodeHTMLEntities(concentracion)
		etiquetadoFormated = encodeHTMLEntities(etiquetado)
		etiConc.add "concentracion", concentracionFormated
		etiConc.add "etiquetado", etiquetadoFormated
		set obtainConcentracionEtiquetado = etiConc
		exit function
	end if


	set obtainConcentracionEtiquetado = etiConc
	exit function
end function
%>