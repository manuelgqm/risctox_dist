<%
function findClasificacionesRd1272(substance, connection)
	Dim clasificacionesRaw : clasificacionesRaw = getClasificacionesRaw(substance)
	dim clasificaciones : clasificaciones = extractClasificaciones(clasificacionesRaw, connection)

	findClasificacionesRd1272 = clasificaciones
end function

function getClasificacionesRaw(substance)
	Dim result : result = Array( _
			substance.item("clasificacion_rd1272_1"), _
			substance.item("clasificacion_rd1272_2"), _
			substance.item("clasificacion_rd1272_3"), _
			substance.item("clasificacion_rd1272_4"), _
			substance.item("clasificacion_rd1272_5"), _
			substance.item("clasificacion_rd1272_6"), _
			substance.item("clasificacion_rd1272_7"), _
			substance.item("clasificacion_rd1272_8"), _
			substance.item("clasificacion_rd1272_9"), _
			substance.item("clasificacion_rd1272_10"), _
			substance.item("clasificacion_rd1272_11"), _
			substance.item("clasificacion_rd1272_12"), _
			substance.item("clasificacion_rd1272_13"), _
			substance.item("clasificacion_rd1272_14"), _
			substance.item("clasificacion_rd1272_15") _
		)

	getClasificacionesRaw = result
End function

function extractClasificaciones(clasificacionesRaw, connection)
	Dim i
	Dim result : result = Array()
	For i = 0 to UBound(clasificacionesRaw)
		if clasificacionesRaw(i) <> "" then
			Set clasificacion = obtainClasificacion(clasificacionesRaw(i), connection)
			result = arrayPushDictionary(result, clasificacion)
		end if
	Next

	extractClasificaciones = result
end function

function obtainClasificacion(clasificacionRaw, connection)
	Dim result : Set result = Server.CreateObject("Scripting.Dictionary")
	
	Dim clasificacionDecomposed : clasificacionDecomposed = getClasificacionDecomposed(clasificacionRaw)
	Dim categoriaPeligroRaw	: categoriaPeligroRaw = clasificacionDecomposed(0)
	Dim categoriaPeligroDecomposed : categoriaPeligroDecomposed = split(categoriaPeligroRaw, ",")
	Dim frase : frase = obtainFrase(clasificacionDecomposed)
	result.add "frase", frase
	result.add "fraseDescription", findFraseHDescription(frase, connection)
	result.add "categoriaPeligro", obtainCategoriaPeligro(categoriaPeligroDecomposed)
	result.add "categoriaPeligroDescription", obtaincategoriaPeligroDescription(categoriaPeligroDecomposed, connection)

	Set obtainClasificacion = result
end function

function getClasificacionDecomposed(clasificacionRaw)
	Dim result : result = split(clasificacionRaw, ";")
	result(0) = trim(result(0))
	result(1) = trim(result(1))
	
	getClasificacionDecomposed = result
end function

function obtainFrase(clasificacionDecomposed)
	Dim frase : frase = clasificacionDecomposed(0)
	if ubound(clasificacionDecomposed)>0 then
		frase = clasificacionDecomposed(1)
	end if
	if frase = "H???" then
		frase = "Gases a presión"
	end if

	obtainFrase = frase
end function

function obtainCategoriaPeligro(categoriaPeligroDecomposed)
	Dim result : result = ""
	if ubound(categoriaPeligroDecomposed) > 0 then
		result = "Cat. " + categoriaPeligroDecomposed(1)
	end if

	obtainCategoriaPeligro = result
end function

function obtaincategoriaPeligroDescription(categoriaPeligroDecomposed, connection)
	Dim result : result = ""
	if ubound(categoriaPeligroDecomposed) < 0 then
		obtaincategoriaPeligroDescription = result
		Exit function
	end if
	result = findCategoriaPeligroDescription(categoriaPeligroDecomposed(0), connection)

	obtaincategoriaPeligroDescription = result
end function

function findFraseHDescription(frase, connection)
	Dim result : result = ""
	' Sustituye "-" por "/" para unificar formato
	frase = replace(frase, "-", "/")
	frase = replace(frase, "*", "")
	Dim sql, objRst
	sql = "SELECT dbo.udf_StripHTML(texto) as texto FROM dn_risc_frases_h WHERE frase = '" & frase & "'"
	set objRst = connection.execute(sql)
	if (objRst.eof) then
		findFraseHDescription = result
		Exit function
	end if
	result = objRst("texto")

	objRst.close()
	set objRst = nothing
	findFraseHDescription = result
end function

function findCategoriaPeligroDescription(categoria, connection)
	Dim result : result = ""
	' Sustituye "-" por "/" para unificar formato
	frase = replace(categoria, "-", "/")
	Dim sql, objRst
	sql = "SELECT texto FROM dn_risc_categorias_peligro WHERE frase = '" & categoria & "'"
	set objRst = connection.execute(sql)
	if (objRst.eof) then
		findFraseHDescription = result
		Exit function
	end if
	result = objRst("texto")

	objRst.close()
	set objRst=nothing
	findCategoriaPeligroDescription = result
end function
%>