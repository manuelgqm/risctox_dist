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
		pictogram.add "imageUrl", getImageUrl("", simbolo)
		Set extractPictograms = pictogram
		exit function
	end if
	pictogram.add "name", simbolo
	pictogram.add "image", simbolosRecodset("imagen").value
	pictogram.add "description", simbolosRecodset("descripcion").value
	pictogram.add "imageUrl", getImageUrl(simbolosRecodset("imagen").value, simbolo)

	simbolosRecodset.close()
	set simbolosRecodset = nothing
	Set extractPictograms = pictogram
end function

function getImageUrl(image, simbolo)
	const PICTOGRAMS_IMAGES_BASE_PATH = "../imagenes/pictogramas/"
	const PELIGRO_IMAGE = "pictograma_peligro.gif"
	const ATENCION_IMAGE = "pictograma_atencion.gif"
	getImageUrl = PICTOGRAMS_IMAGES_BASE_PATH + image
	if simbolo = "Peligro" then getImageUrl = PICTOGRAMS_IMAGES_BASE_PATH + PELIGRO_IMAGE
	if simbolo = "Atención" then getImageUrl = PICTOGRAMS_IMAGES_BASE_PATH + ATENCION_IMAGE
end function
%>