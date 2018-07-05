<%
function obtainNotasRd1272(byVal notasSrz, lang, connection)
	obtainNotasRd1272 = Array()
	if is_empty(notasSrz) then _
		exit function
	notasSrz = removeTailSeparator(notasSrz, ".")
	dim notas : notas = split(notasSrz, ".")
	notas = arrayWrapItems(notas, "R.1272-", "")
	dim notasSrzFormatted : notasSrzFormatted = join(notas, ",")
	obtainNotasRd1272 = clearKeys("R.1272-", obtainDefinitions(notasSrzFormatted, lang, connection))
end function

function clearKeys(str, notas)
	dim i, current
	for i = 0 to Ubound(notas)
		current = notas(i)("key")
		notas(i)("key") = replace(current, str, "")
	next
	clearKeys = notas
end function
%>
