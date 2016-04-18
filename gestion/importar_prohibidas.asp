<!--#include file="dn_conexion.asp"-->

<%

FUNCTION unQuote(s)
  pos = Instr(s, "'")
  While pos > 0 
    s = Mid(s,1,pos) & "'" & Mid(s,pos+1)
    pos = InStr(pos+2, s, "'")
  Wend
  pos = Instr(s, """")
  While pos > 0 
    s = Mid(s,1,pos-1) & "''" & Mid(s,pos+1)
    pos = InStr(pos+2, s, """")
  Wend
  unQuote = Trim(s)
END FUNCTION

function arreglar(x)
	if isnull(x) then x=""
	x_ = replace(x,"'","´")
	arreglar = unquote(x_)
end function

'PROHIBIDAS
'sql = "select * from temporal_prohibidas where prohibido is null"
sql = "select * from temporal_prohibidas where (NOT(prohibido IS NULL))"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
		
response.write "<table border=1 cellspacing=1 cellspadding=1>"
while not objr.eof

	cas = objr("cas")
	if isnull(cas) then cas=""
	cas = trim(cas)
	cas = replace(cas,chr(10),"")
	cas = replace(cas,chr(13),"")
	ce = objr("einecs")
	limite_exencion = objr("limite_exencion")
	limitaciones = objr("limitaciones")
	fuente = objr("fuente")

	comentario = ""
	if not isnull(limite_exencion) then comentario = comentario & "<p><b>Límite exención</b>: "&arreglar(limite_exencion)&"</p>"
	if not isnull(limitaciones) then comentario = comentario & "<p><b>Limitaciones</b>: "&arreglar(limitaciones)&"</p>"
	if not isnull(fuente) then comentario = comentario & "<p><b>Fuente</b>: "&arreglar(fuente)&"</p>"
	
	if trim(objr("cas"))<>"" then
		sql2 = "select id, comentarios from dn_risc_sustancias where num_cas='"&cas&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		if not objr2.eof then
			id_sustancia = objr2("id")
			sql3 = "INSERT INTO dn_risc_sustancias_prohibidas (id_sustancia,comentario_prohibida) values ('"&id_sustancia&"','"&comentario&"')"
			objconn1.execute(sql3)
			response.write "<tr><td>"&cas&"</td></tr>"
		end if		
	end if
	
	objr.movenext
wend
response.write "</table>"
%>