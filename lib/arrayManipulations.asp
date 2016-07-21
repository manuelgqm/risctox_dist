<%
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
%>