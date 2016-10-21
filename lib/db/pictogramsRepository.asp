<%
function findPictogramasRd1272(simbolosRdString, connection)
	Dim simbolosRd, i, simbolo, pictogram
	Dim pictograms : pictograms = Array()
	
	simbolosRdString = removeTailSeparator(simbolosRdString, ",")
	if isNull(simbolosRdString) then
		findPictogramasRd1272 = pictograms
		exit function
	end if
	simbolosRd = split(simbolosRdString, ",")
	for i = 0 to ubound(simbolosRd)
		simbolo = trim(simbolosRd(i))
		Set pictogram = extractPictograms(simbolo, connection)
		pictograms = arrayPushDictionary(pictograms, pictogram)
	next
	
	findPictogramasRd1272 = pictograms
End function

function extractPictograms(simbolo, connection)
	dim sql, simbolosRecodset
	dim pictogram : Set pictogram = Server.CreateObject("Scripting.Dictionary")
	simbolo = trim(simbolo)
	
	sql = "SELECT imagen, descripcion FROM dn_simbolos WHERE simbolo='" & simbolo & "'"
	set simbolosRecodset = connection.execute(sql)
	if simbolosRecodset.eof then
		pictogram.add "name", simbolo
		pictogram.add "image", ""
		pictogram.add "description", simbolo
		Set extractPictograms = pictogram
		exit function
	end if
	pictogram.add "name", simbolo
	pictogram.add "image", simbolosRecodset("imagen").value
	pictogram.add "description", simbolosRecodset("descripcion").value

	simbolosRecodset.close()
	set simbolosRecodset = nothing
	Set extractPictograms = pictogram
end function
%>