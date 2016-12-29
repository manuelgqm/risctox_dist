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
	dim i, fraseH
	For i = 0 to UBound(frasesHRaw)
		if not isEmpty(frasesHRaw(i)) then
			Set fraseH = obtainFraseH(frasesHRaw(i), connection)
			extractFrasesH = arrayPushDictionary(extractFrasesH, fraseH)
		end if
	Next
end function

function obtainFraseH(fraseHRaw, connection)
	set obtainFraseH = Server.CreateObject("Scripting.Dictionary")
	dim fraseHDecomposed : fraseHDecomposed = decomposeFraseH(fraseHRaw)
	dim peligroRaw	: peligroRaw = fraseHDecomposed(0)
	dim peligroDecomposed : peligroDecomposed = split(peligroRaw, ",")
	dim frase : frase = extractFraseH(fraseHDecomposed)
	obtainFraseH.add "fraseH", frase
	obtainFraseH.add "fraseHDescription", findFraseHDescription(frase, connection)
	obtainFraseH.add "peligro", obtainPeligro(peligroDecomposed)
	obtainFraseH.add "peligroDescription", obtainpeligroDescription(peligroDecomposed, connection)
end function

function decomposeFraseH(fraseHRaw)
	Dim result : result = split(fraseHRaw, ";")
	result(0) = trim(result(0))
	result(1) = trim(result(1))
	
	decomposeFraseH = result
end function

function extractFraseH(fraseHDecomposed)
	Dim fraseH : fraseH = fraseHDecomposed(0)
	if ubound(fraseHDecomposed)>0 then
		fraseH = fraseHDecomposed(1)
	end if
	if fraseH = "H???" then
		fraseH = "Gases a presión"
	end if

	extractFraseH = fraseH
end function

function obtainPeligro(peligroDecomposed)
	Dim result : result = ""
	if ubound(peligroDecomposed) > 0 then
		result = "Cat. " + peligroDecomposed(1)
	end if

	obtainpeligro = result
end function

function obtainpeligroDescription(peligroDecomposed, connection)
	Dim result : result = ""
	if ubound(peligroDecomposed) < 0 then
		obtainpeligroDescription = result
		Exit function
	end if
	result = findpeligroDescription(peligroDecomposed(0), connection)

	obtainpeligroDescription = result
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

function findpeligroDescription(categoria, connection)
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
	findpeligroDescription = result
end function
%>