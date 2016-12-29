<%
function findfrasesH(substance, connection)
	dim frasesH : frasesH = Array( _
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

	findfrasesH = extractFrasesH(frasesH, connection)
end function

function findFrasesR(frasesRSrz, connection)
	findFrasesR = Array()
	dim frasesR : frasesR = split(frasesRSrz, ", ")
	dim i, fraseR
	for i = 0 to Ubound(frasesR)
		set fraseR = Server.CreateObject("Scripting.Dictionary")
		fraseR.add "name", frasesR(i)
		fraseR.add "description", findFraseRDescription(frasesR(i), connection)
		findFrasesR = arrayPushDictionary(findFrasesR, fraseR)
	next
end function

function findFrasesS(byVal frasesSSrz, connection)
	findFrasesS = Array()
	frasesSSrz = replace( _
		replace( _
			replace( frasesSSrz, "(", "") _
		, ")", "") _
	, "S:", "")
	dim frasesS : frasesS = split(frasesSSrz, "-")
	dim i, fraseS, fraseSName
	for i = 0 to Ubound(frasesS)
		set fraseS = Server.CreateObject("Scripting.Dictionary")
		fraseSName = "S" & trim(frasesS(i))
		fraseS.add "name", fraseSName
		fraseS.add "description", findFraseSDescription(fraseSName, connection)
		findFrasesS = arrayPushDictionary(findFrasesS, fraseS)
	next
end function

function extractFrasesH(frasesHRaw, connection)
	extractFrasesH = Array()
	Dim i
	For i = 0 to UBound(frasesHRaw)
		if frasesHRaw(i) <> "" then
			Set clasificacion = obtainClasificacion(frasesHRaw(i), connection)
			extractFrasesH = arrayPushDictionary(extractFrasesH, clasificacion)
		end if
	Next
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

function findFraseRDescription(byVal frase, connection)
	findFraseRDescription = ""
	frase = replace(frase, "-", "/")
	dim sql : sql = "SELECT texto as texto FROM dn_risc_frases_r WHERE frase = '" & frase & "'"
	dim recordset : set recordset = connection.execute(sql)
	if recordset.eof then _
		exit function
	findFraseRDescription = recordset("texto").value
	recordset.close
	set recordset = nothing
end function

function findFraseSDescription(byVal frase, connection)
	findFraseSDescription = ""
	frase = replace(frase, "-", "/")
	dim sql : sql = "SELECT texto as texto FROM dn_risc_frases_s WHERE frase = '" & frase & "'"
	dim recordset : set recordset = connection.execute(sql)
	if recordset.eof then _
		exit function
	findFraseSDescription = recordset("texto").value
	recordset.close
	set recordset = nothing
end function

function findCategoriaPeligroDescription(categoria, connection)
	Dim result : result = ""
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