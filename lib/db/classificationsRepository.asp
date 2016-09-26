<%
function findClasificacionesRd1272(substance)
	dim clasificacionesRd1272 : Array()
	clasificacionesRd1272 = extractClasificacionesRd1272(substance)

	findClasificacionesRd1272 = clasificacionesRd1272
end function

function extractClasificacionesRd1272(substance)
	Dim clasificacionesRaw : clasificacionesRaw = Array( _
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

	Dim i
	Dim result : result = Array()
	For i = 0 to UBound(clasificacionesRaw)
		if clasificacionesRaw(i) <> "" then
			Set clasificacion = obtainClasificacionRd(clasificacionesRaw(i))
			result = arrayPushDictionary(result, clasificacion)
		end if
	Next

	extractClasificacionesRd1272 = result
end function

function obtainClasificacionRd(clasificacionRaw)
	Dim result : Set result = Server.CreateObject("Scripting.Dictionary")
	
	Dim clasificacionDecomposed : clasificacionDecomposed = split(clasificacionRaw, ";")
	clasificacionDecomposed(0) = trim(clasificacionDecomposed(0))
	clasificacionDecomposed(1) = trim(clasificacionDecomposed(1))
	Dim categoriaPeligro : categoriaPeligro = clasificacionDecomposed(0)
	Dim categoriaPeligroDecomposed : categoriaPeligroDecomposed = split(categoriaPeligro, ",")
	categoriaPeligro = obtainCategoriaPeligro(categoriaPeligroDecomposed)
	Dim categoriaPeligroDescription : categoriaPeligroDescription = obtaincategoriaPeligroDescription(categoriaPeligroDecomposed)
	Dim frase : frase = obtainFrase(clasificacionDecomposed)
	Dim fraseDescription : fraseDescription = obtainFraseHDescription(frase)
	result.add "frase", frase
	result.add "fraseDescription", fraseDescription
	result.add "categoriaPeligro", categoriaPeligro
	result.add "categoriaPeligroDescription", categoriaPeligroDescription

	Set obtainClasificacionRd = result
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

function obtainFraseHDescription(frase)
	' Sustituye "-" por "/" para unificar formato
	frase = replace(frase, "-", "/")
	frase = replace(frase, "*", "")

	sql = "SELECT dbo.udf_StripHTML(texto) as texto FROM dn_risc_frases_h WHERE frase = '" & frase & "'"
	set objRst = objConnection2.execute(sql)

	if (objRst.eof) then
		descripcion = ""
	else
		descripcion = objRst("texto")
	end if

	objRst.close()
	set objRst = nothing

	obtainFraseHDescription = descripcion
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