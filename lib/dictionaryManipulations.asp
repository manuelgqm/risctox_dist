<%
function isDictionaryEmpty(dictionary)
	isDictionaryEmpty = true
	if dictionary.count = 0 then exit function
	dim dictItems : dictItems = dictionary.items
	dim i, dictItem
	for i = 0 to ubound(dictItems)
		if hasValue(dictItems(i)) then 
			isDictionaryEmpty = false
			exit function
		end if
	next
end function

function hasValue(var)
	hasValue = false
	select case varType(var)
		case vbString:
			if len(var) > 0 then hasValue = true
		case vbArray:
			if ubound(var) > -1 then hasValue = true
		case 8204:
			if ubound(var) > -1 then hasValue = true
		case vbBoolean:
			hasValue = true
	end select
end function
%>