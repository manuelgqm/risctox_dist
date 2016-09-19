<%
function findPictogramasRd1272(simbolosRdString, connection)
	Dim simbolosRd, i, simbolo, pictogram
	Dim pictograms : pictograms = Array()
	
	simbolosRdString = removeTailSeparator(simbolosRdString, ",")
	simbolosRd = split(simbolosRdString, ",")
	for i = 0 to ubound(simbolosRd)
		simbolo = trim(simbolosRd(i))
		Set pictogram = findPictogram(simbolo, connection)
		pictograms = arrayPushDictionary(pictograms, pictogram)
	next
	
	findPictogramasRd1272 = pictograms
End function

function findPictogram(simbolo, connection)
	dim sql, simbolosRecodset
	dim pictogram : Set pictogram = Server.CreateObject("Scripting.Dictionary")
	simbolo = trim(simbolo)
	
	sql = "SELECT imagen, descripcion FROM dn_simbolos WHERE simbolo='" & simbolo & "' and simbolo like 'GHS%'"
	set simbolosRecodset = connection.execute(sql)
	if simbolosRecodset.eof then
		pictogram.add "name", simbolo
		pictogram.add "image", ""
		pictogram.add "description", simbolo
		Set findPictogram = pictogram
		exit function
	end if
	pictogram.add "name", simbolo
	pictogram.add "image", simbolosRecodset("imagen").value
	pictogram.add "description", simbolosRecodset("descripcion").value

	simbolosRecodset.close()
	set simbolosRecodset = nothing
	Set findPictogram = pictogram
end function

function decidePictogramDescription(image, description, simbolo)
	if image = "" Then
		decidePictogramDescription = simbolo
		Exit function
	end if

	decidePictogramDescription = description
end function

function removeTailSeparator(str, separator)
	dim result

	result = str
	if Right(str, 1) = separator then
		result = Left(str, Len(str) - 1)
	end if

	removeTailSeparator = result
end function
%>