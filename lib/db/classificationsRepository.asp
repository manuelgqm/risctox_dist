<%
function findClasificacionesRd1272(substance, connection)
	Dim clasificacionesRd1272Raw : clasificacionesRd1272Raw = getClasificacionesRd1272Raw(substance)
	dim clasificacionesRd1272 : Array()
	clasificacionesRd1272 = extractClasificacionesRd1272(clasificacionesRd1272Raw, connection)

	findClasificacionesRd1272 = clasificacionesRd1272
end function

function getClasificacionesRd1272Raw(substance)
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

	getClasificacionesRd1272Raw = result
End function

function extractClasificacionesRd1272(clasificacionesRd1272Raw, connection)
	Dim i
	Dim result : result = Array()
	For i = 0 to UBound(clasificacionesRd1272Raw)
		if clasificacionesRd1272Raw(i) <> "" then
			Set clasificacion = obtainClasificacionRd(clasificacionesRd1272Raw(i), connection)
			result = arrayPushDictionary(result, clasificacion)
		end if
	Next

	extractClasificacionesRd1272 = result
end function

function obtainClasificacionRd(clasificacionRaw, connection)
	Dim result : Set result = Server.CreateObject("Scripting.Dictionary")
	
	Dim clasificacionDecomposed : clasificacionDecomposed = getClasificacionDecomposed(clasificacionRaw)
	Dim categoriaPeligroRaw	: categoriaPeligroRaw = clasificacionDecomposed(0)
	Dim categoriaPeligroDecomposed : categoriaPeligroDecomposed = split(categoriaPeligroRaw, ",")
	Dim frase : frase = obtainFrase(clasificacionDecomposed)
	result.add "frase", frase
	result.add "fraseDescription", findFraseHDescription(frase, connection)
	result.add "categoriaPeligro", obtainCategoriaPeligro(categoriaPeligroDecomposed)
	result.add "categoriaPeligroDescription", obtaincategoriaPeligroDescription(categoriaPeligroDecomposed)

	Set obtainClasificacionRd = result
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

function obtaincategoriaPeligroDescription(categoriaPeligroDecomposed)
	Dim result : result = ""
	if ubound(categoriaPeligroDecomposed) < 0 then
		obtaincategoriaPeligroDescription = result
		Exit function
	end if
	result = findCategoriaPeligroDescription(categoriaPeligroDecomposed(0))

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
		result = ""
		findFraseHDescription = result	
	end if

	result = objRst("texto")

	objRst.close()
	set objRst = nothing

	findFraseHDescription = result
end function

function findCategoriaPeligroDescription(categoria)
	' Sustituye "-" por "/" para unificar formato
	frase = replace(categoria, "-", "/")

	sql = "SELECT texto FROM dn_risc_categorias_peligro WHERE frase='" & categoria & "'"
	set objRst=objConnection2.execute(sql)

	if (objRst.eof) then
		descripcion = ""
	else
		descripcion = objRst("texto").value
	end if

	objRst.close()
	set objRst=nothing

	findCategoriaPeligroDescription = descripcion
end function
%>