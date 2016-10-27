<%
function removeDuplicates(byval inputStringList, byval separatorChar)
	dim outputStringList : outputStringList = ""
	
	if inputStringList = "" then
		removeDuplicates = ""
		exit function
	end if

	dim arrayList
	arrayList = split(inputStringList, separatorChar)
	
	dim i, actualItem
	for i = 0 to ubound(arrayList)
	actualItem = arrayList(i)
	if not stringContains(outputStringList, actualItem) then
		outputStringList = outputStringList & actualItem & addListSeparator(i, ubound(arrayList), separatorChar)
	end if
	next
	removeDuplicates = outputStringList
end function

function stringContains(container, content)
	stringContains = false
	if container = "" or content = "" then
		exit function
	end if
	if instr(lcase(container), lcase(content)) > 0 then
		stringContains = true
	end if
end function

function addListSeparator(currentIndex, lastIndex, separatorChar)
	dim calculatedSeparatorChar : calculatedSeparatorChar = ""
	if not currentIndex + 1 > lastIndex then
	calculatedSeparatorChar = separatorChar
	end if

	addListSeparator = calculatedSeparatorChar
end function

function removeTailSeparator(str, separator)
	dim result

	result = str
	if Right(str, 1) = separator then
	result = Left(str, Len(str) - 1)
	end if

	removeTailSeparator = result
end function

function encodeHTMLEntities(byval cadena)
	if (isNull(cadena)) then
		encodeHTMLEntities = ""
		exit function
	end if
	cadena = trim(cadena)
	cadena = Replace(cadena,"´","'")
	cadena = Replace(cadena,"'","&#39;")
	cadena = Replace(cadena,"’","&#39;")
	cadena = Replace(cadena,"""","&#34;")
	cadena = Replace(cadena,"%","&#37;")
	cadena = Replace(cadena,"[","&#91;")
	cadena = Replace(cadena,"]","&#93;")
	cadena = Replace(cadena,"<", "&lt;")
	cadena = Replace(cadena,">", "&gt;")
	
	encodeHTMLEntities = cadena
end function
%>