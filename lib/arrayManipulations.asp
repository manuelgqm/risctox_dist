<%
' depends stringManipulatios.asp'
function arrayPushDictionary(arrayParameter, dictionaryParameter) 
	dim result : result = arrayParameter
	if vartype(dictionaryParameter) <> 9 or not isArray(arrayParameter) Then
		arrayPushDictionary = result
		exit function
	end if

	dim newId : newId = ubound(arrayParameter) + 1
	redim preserve result(newId) 
	set result(newId) = dictionaryParameter

	arrayPushDictionary = result
end function

function arraySerialize(arr)
	dim result : result = ""
	dim i

	if not isArray(arr) Then
		arraySerialize = result
		exit function
	end if
	for i = 0 to uBound(arr)
		result = result & arr(i) & addListSeparator(i, uBound(arr), ", ")
	next
	
	arraySerialize = result
end function
%>