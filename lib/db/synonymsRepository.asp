<%
function obtainSynonyms(byval id_sustancia, connection)
	dim objRst, counter
	counter = 0

	sql="SELECT dn_risc_sinonimos.nombre AS sinonimo, dn_risc_sustancias.nombre FROM dn_risc_sinonimos INNER JOIN dn_risc_sustancias ON dn_risc_sinonimos.id_sustancia = dn_risc_sustancias.id WHERE dn_risc_sinonimos.id_sustancia="& id_sustancia &" AND dn_risc_sinonimos.nombre <> dn_risc_sustancias.nombre ORDER BY dn_risc_sinonimos.nombre"
	set objRst = Server.CreateObject("ADODB.recordset")
	objRst.CursorType = 3
	objRst.Open sql, connection
	
	if (objRst.eof) then obtainSynonyms = vbNull
	
	redim synonyms(cint(objRst.RecordCount)-1)
	do while (not objRst.eof)
		synonyms(counter) = objRst("sinonimo")
		counter  = counter + 1
		objRst.movenext
	loop
	
	objRst.close()
	set objRst=nothing

	obtainSynonyms = synonyms
end function
%>