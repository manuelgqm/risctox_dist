<%
function obtainConcentracionEtiquetadoRd1272(substance)
	Dim etiConcList : etiConcList = Array()
	Dim etiConc : set etiConc = Server.CreateObject("Scripting.Dictionary")
	if substance.item("conc_rd1272_1") = "" and substance.item("conc_rd1272_2") = "" and substance.Item("eti_conc_rd1272_1") = "" Then
		etiConc.add "concentracion", ""
		etiConcList = arrayPushDictionary(etiConcList, etiConc)
		Exit function
	end if
	etiConcList = extractEtiConcList(substance)

	obtainConcentracionEtiquetadoRd1272 = etiConcList
end function

function obtainConcentracionEtiquetadoRd363(substance)
	obtainConcentracionEtiquetadoRd363 = Array()
	dim etiquetas : etiquetas = Array( _
		"eti_conc_1", _
		"eti_conc_2", _
		"eti_conc_3", _
		"eti_conc_4", _
		"eti_conc_5", _
		"eti_conc_6", _
		"eti_conc_7", _
		"eti_conc_8", _
		"eti_conc_9", _
		"eti_conc_10", _
		"eti_conc_11", _
		"eti_conc_12", _
		"eti_conc_13", _
		"eti_conc_14", _
		"eti_conc_15" _
	)
	dim concentraciones : concentraciones = Array( _
		"conc_1", _
		"conc_2", _
		"conc_3", _
		"conc_4", _
		"conc_5", _
		"conc_6", _
		"conc_7", _
		"conc_8", _
		"conc_9", _
		"conc_10", _
		"conc_11", _
		"conc_12", _
		"conc_13", _
		"conc_14", _
		"conc_15" _
	)
	dim i, currentEtiqueta, currentConcentracion, etiConc
	for i = 0 to Ubound(etiquetas)
		currentEtiqueta = getField(substance, etiquetas(i))
		currentConcentracion = getField(substance, concentraciones(i))
		set etiConc = Server.CreateObject("Scripting.Dictionary")
		if not isEmpty(currentEtiqueta) and not isEmpty(currentConcentracion) then
			etiConc.add "concentracion", currentConcentracion
			etiConc.add "etiquetado", currentEtiqueta
			obtainConcentracionEtiquetadoRd363 = arrayPushDictionary(obtainConcentracionEtiquetadoRd363, etiConc)
		end if
	next
end function

function getField(substance, byVal fieldName)
	getField = ""
	dim value : value = substance(fieldName)
	if isEmpty(value) then _
		exit function
	getField = encodeHTMLEntities( replace(value, ":", "") )
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
		dim currentEtiquetado : currentEtiquetado = substance.item(etiquetas(i))
		dim currentConcentracion : currentConcentracion = substance.item(concentraciones(i))
		set etiConc = obtainConcentracionEtiquetado(currentEtiquetado, currentConcentracion)
		if etiConc.count > 0 then
			result = arrayPushDictionary(result, etiConc)
		end if
		if currentEtiquetado = "*" then
			extractEtiConcList = result
			exit function
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

	if concentracion <> "" or etiquetado <> "" Then
		if not is_empty(concentacion) then
			concentracionFormated = replace(concentracion, ":", "")
			concentracionFormated = encodeHTMLEntities(concentracion)
			etiConc.add "concentracion", concentracionFormated
		end if
		if not is_empty(etiquetado) then
			etiquetadoFormated = encodeHTMLEntities(etiquetado)
			etiConc.add "etiquetado", etiquetadoFormated
		end if
		set obtainConcentracionEtiquetado = etiConc
		exit function
	end if

	set obtainConcentracionEtiquetado = etiConc
	exit function
end function
%>
