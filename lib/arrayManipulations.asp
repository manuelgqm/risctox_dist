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

function arrayWrapItems(arr, prepend, append)
	dim result : result = Array()
	if not isArray(arr) Then
		arrayWrapItems = result
		exit function
	end if
	dim i, itemConverted
	for i = 0 to uBound(arr)
		itemConverted = prepend & arr(i) & append
		arrayPush result, itemConverted
	next
	
	arrayWrapItems = result
end function

Sub arrayPush(byRef arrayParameter, valueParameter) 
	redim preserve arrayParameter(uBound(arrayParameter) + 1)
	arrayParameter(uBound(arrayParameter)) = valueParameter
End Sub

function inArray(element, arrayParameter)
	inArray = false
	if not isArray(arrayParameter) or element = "" then
		exit function
	end if
	dim i
	For i = 0 To Ubound(arrayParameter)
		If Trim(arrayParameter(i)) = Trim(element) Then 
			inArray = true
			Exit Function
		end if
	Next
End Function

function anyElementInArray(sources, targets)
	anyElementInArray = false
	if not isArray(sources) then
		exit function
	end if
	dim i
	for i = 0 to Ubound(sources)
		if inArray(sources(i), targets) then 
			anyElementInArray = true
			exit function
		end if
	next
end function
%>