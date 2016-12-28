<%
function findPictograms(byVal simbolosSrz, connection)
	findPictograms = Array()
	if isNull(simbolosSrz) then _
		exit function
	simbolosSrz = replace(simbolosSrz, ";", ",")
	simbolosSrz = removeTailSeparator(simbolosSrz, ",")
	dim simbolos
	simbolos = split(simbolosSrz, ",")
	dim i, simbolo, pictogram
	for i = 0 to ubound(simbolos)
		simbolo = trim(simbolos(i))
		Set pictogram = extractPictograms(simbolo, connection)
		findPictograms = arrayPushDictionary(findPictograms, pictogram)
	next
End function

function extractPictograms(simbolo, connection)
	dim sql
	dim pictogram : Set pictogram = Server.CreateObject("Scripting.Dictionary")
	simbolo = trim(simbolo)
	dim simbolosRecodset : set simbolosRecodset = findSimbolos(simbolo, connection)
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

function findSimbolos(simbolo, connection)
	dim sql : sql = "SELECT imagen, descripcion FROM dn_simbolos WHERE simbolo='" & simbolo & "'"
	set findSimbolos = connection.execute(sql)
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