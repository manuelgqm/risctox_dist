<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
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


response.End()
sql = "select * from temporal_anexo_reach"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table border=1>"
while not objr.eof
	if trim(objr("cas"))<>"" then
		sql2 = "select id from dn_risc_sustancias where num_cas='"&trim(objr("cas"))&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		response.write "<tr>"
		if objr2.eof then
			response.write "<td>NO EXISTE CAS "&objr("cas")&"</td>"
		else
			'Miro si esa sustantia tiene uso
			sql3="select id, id_uso from dn_risc_sustancias_por_usos where id_sustancia="&objr2("id")& "and id_uso="&trim(objr("uso"))
			Set objr3 = Server.CreateObject ("ADODB.Recordset")
			objr3.Open sql3,objconn1,adOpenKeyset
			if not objr3.eof then
				'response.write "<td>YA LO TIENE "&objr("nombre")&"</td><td>"&objr3("id")&"SI </td><td>"&objr3("id_uso")&"</td>"
				sql4="update dn_risc_sustancias_por_usos set toxico=0, anexo_reach=1 where id="&objr3("id")
				objconn1.execute(sql4)
				response.write "<td>"&sql4&"</td>"
			else	
				sql4="insert into dn_risc_sustancias_por_usos(id_sustancia, id_uso, toxico, anexo_reach) values("&objr2("id")&","&objr("uso")&",0,1)"
				objconn1.execute(sql4)
				response.write "<td>"&sql4&"</td>"
			end if
			
			
		end if
		response.write "</tr>"
	end if
	objr.movenext
wend
response.write "</table>"
response.end()


response.End()
'PROHIBIDAS-RESTRINGIDAS
sql = "select * from temporal_prohibidas where prohibido is null"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table border=1 cellspacing=1 cellspadding=1>"
while not objr.eof
	if trim(objr("cas"))<>"" then
			
		'Compruebo si ese numero CAS está en la BBDD
		sql2 = "select id, comentarios from dn_risc_sustancias where num_cas='"&replace(replace(trim(objr("cas")),chr(13),""),chr(10),"")&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		if not objr2.eof then
			'response.write "<td>SI</td>"
			' Añado a la tabla dn_risc_sensibilizantes_reach
			if trim(objr("fuente"))<>"" then
				sql3="update dn_risc_sustancias set sustancia_restringida=1, comentarios='"&objr2("comentarios")&objr("limitaciones")&" FUENTE: "&objr("fuente")&"' where id="&objr2("id")
				
			else
			
				sql3="update dn_risc_sustancias set sustancia_restringida=1, comentarios='"&objr("limitaciones")&"' where id="&objr2("id")
				
			end if
				
				objconn1.execute(sql3)
				
		else
			
			'response.write "<tr>"
			'response.write "<td>"&objr("nombre")&"</td>"
			'response.write "<td>"&objr("cas")&"</td>"
			'response.write "<td>"&sql2&"</td>"
			'response.write "</tr>"
			
			'response.write "<td>NO</td>"
			'sql_i = "insert into dn_risc_sustancias (nombre,nombre_ing,num_cas,negra,sustancia_prohibida,sustancia_restringida) values('"&trim(unQuote(objr("nombre")))&" (nombre en ingles)"&"','"&trim(unQuote(objr("nombre")))&"','"&trim(objr("cas"))&"',0,0,0)"
			'response.write "<td>"&sql_i&"</td>"
			'objconn1.execute(sql_i)

		end if
	
	else
		'response.write "<tr><td>SIN NUMERO CAS->:"&objr("nombre")&"</td><td>"&objr("einecs")&"</td></tr>"
	
	end if
	objr.movenext
wend
response.write "</table>"

response.end()
'SENSIBILIZANTES REACH
sql = "select * from temporal_sensibilizantes_reach order by nombre asc"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table>"
while not objr.eof
	if trim(objr("cas"))<>"" then
		response.write "<tr>"
		response.write "<td>"&objr("cas")&"</td>"	
		'Compruebo si ese numero CAS está en la BBDD
		sql2 = "select id from dn_risc_sustancias where num_cas='"&trim(objr("cas"))&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		if not objr2.eof then
			response.write "<td>SI</td>"
			' Añado a la tabla dn_risc_sensibilizantes_reach
			sql3="insert into dn_risc_sensibilizantes_reach(id_sustancia) values('"&objr2("id")&"')"
			response.write "<td>"&sql3&"</td>"
			objconn1.execute(sql3)
		else
			response.write "<td>NO</td>"
			sql_i = "insert into dn_risc_sustancias (nombre,nombre_ing,num_cas,negra,sustancia_prohibida,sustancia_restringida) values('"&trim(unQuote(objr("nombre")))&" (nombre en ingles)"&"','"&trim(unQuote(objr("nombre")))&"','"&trim(objr("cas"))&"',0,0,0)"
			response.write "<td>"&sql_i&"</td>"
			'objconn1.execute(sql_i)

		end if

		response.write "</tr>"
			
	
	else
		response.write "<tr><td colspan=2>SIN NUMERO CAS->GRUPO:"&objr("nombre")&"</td></tr>"
	
	end if
	objr.movenext
wend
response.write "</table>"


response.end()
'SUSTANCIAS USOS NUEVAS REVISADAS
sql = "select * from temporal_sustancias_usos_nuevas_ya_revisadas order by codigo_uso asc"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table>"
while not objr.eof
	if trim(objr("cas"))<>"" then
		response.write "<tr>"
		response.write "<td>"&objr("uso")&"</td>"	
		'Compruebo si ese numero CAS está en la BBDD
		sql2 = "select id from dn_risc_sustancias where num_cas='"&trim(objr("cas"))&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		if not objr2.eof then
			response.write "<td>"&objr("cas")&"</td>"
			response.write "<td>SI</td>"'Compruebo si ya está tb en la dn_risc_sustancias_por_usos
			'Ahora compruebo si ya está en la lista de usos
			sql3="select id from dn_risc_sustancias_por_usos where id_uso="&trim(objr("codigo_uso"))&" and id_sustancia="&objr2("id")
			Set objr3 = Server.CreateObject ("ADODB.Recordset")
			objr3.Open sql3,objconn1,adOpenKeyset
			if not objr3.eof then
				response.write "<td>ASIGNADA</td>"
			else
				sql_i = "insert into dn_risc_sustancias_por_usos(id_sustancia,id_uso,toxico) values('"&objr2("id")&"','"&objr("codigo_uso")&"',1)"
				response.write "<td>"&sql_i&"</td>"
				
				'objconn1.execute(sql_i)
				'response.write "<td>"&sql_i&"</td>"
			end if
			
			
		else
			response.write "<td>NO "&objr("cas")&"</td>"
			sql_i = "insert into dn_risc_sustancias (nombre,nombre_ing,num_cas,negra,sustancia_prohibida,sustancia_restringida) values('"&trim(unQuote(objr("nombre")))&" (nombre en ingles)"&"','"&trim(unQuote(objr("nombre")))&"','"&trim(objr("cas"))&"',0,0,0)"
			response.write "<td>"&sql_i&"</td>"
			'objconn1.execute(sql_i)
			'Meto esa sustancia en la BBDD
			'sql_i = "insert into temporal_usos_nuevos values('"&unQuote(trim(objr("codigo_uso")))&"','"&unQuote(trim(objr("uso")))&"','"&unQuote(trim(objr("cas")))&"','"&unQuote(trim(objr("nombre")))&"')"
			'objconn1.execute(sql_i)
			'response.write "<tr>"
			'response.write "<td>"&objr("cas")&"</td>"
			'response.write "<td>NO</td>"
			'response.write "<td>"&objr("nombre")&"</td>"
			'response.write "<td>"&sql_i&"</td>"
			'response.write "</tr>"	
		end if

		response.write "</tr>"
			
				
	end if
	objr.movenext
wend
response.write "</table>"

response.End()
'Matriz de autoevaluación
sql = "select * from temporal_matriz"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
while not objr.eof
	sql2="insert into dn_auto_evaluacion (razon, aguda, cronica, ecotoxicidad, fuego, exposicion, proceso)"
	sql2 = sql2 & " values('"&objr("frase")&"','"&objr("toxicidad_aguda")&"','"&objr("toxicidad_cronica")&"','"&objr("peligros_medio_ambiente")&"','"&objr("fuego_y_explosion")&"','"&objr("exposicion")&"','"&objr("proceso")&"')"
	response.write sql2
	response.write "<br>"
	objconn1.execute(sql2)
	objr.movenext()
wend




response.end()
'Para copìar los enlaces
sql = "select * from temporal_cops"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table border=1>"
while not objr.eof
	if trim(objr("cas"))<>"" then
		sql2 = "select id from dn_risc_sustancias where num_cas='"&trim(objr("cas"))&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		response.write "<tr>"
		if objr2.eof then
			response.write "<td>"&objr("nombre")&"</td><td>"&trim(objr("cas"))&"</td>"
		else
			'Miro a ver si está en la tabla dn_risc_sustancias_mama_cop
			sql3 = "select id, cop from dn_risc_sustancias_mama_cop where id_sustancia="&objr2("id")&" and cancer_mama=0"
			Set objr3 = Server.CreateObject ("ADODB.Recordset")
			objr3.Open sql3,objconn1,adOpenKeyset
			if objr3.eof then
				response.write "<td>"&objr2("id")&"</td><td>NUEVO ("&objr("anexo")&")</td>"	
			else
				sql4="update dn_risc_sustancias_mama_cop set enlace_cop='"&objr("enlace")&"' where id="&objr3("id")
				objconn1.execute(sql4)
				response.write "<td>"&objr2("id")&"</td><td>"&sql4&"</td>"	
				
			end if
			
			
		end if
		response.write "</tr>"
	end if
	objr.movenext
wend
response.write "</table>"


response.End()
sql = "select * from temporal_cops"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table border=1>"
while not objr.eof
	if trim(objr("cas"))<>"" then
		sql2 = "select id from dn_risc_sustancias where num_cas='"&trim(objr("cas"))&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		response.write "<tr>"
		if objr2.eof then
			response.write "<td>"&objr("nombre")&"</td><td>"&trim(objr("cas"))&"</td>"
		else
			'Miro a ver si está en la tabla dn_risc_sustancias_mama_cop
			sql3 = "select id, cop from dn_risc_sustancias_mama_cop where id_sustancia="&objr2("id")&" and cancer_mama=0"
			Set objr3 = Server.CreateObject ("ADODB.Recordset")
			objr3.Open sql3,objconn1,adOpenKeyset
			if objr3.eof then
				response.write "<td>"&objr2("id")&"</td><td>NUEVO ("&objr("anexo")&")</td>"	
			else
				if (objr3("cop")="")then
					'response.write "<td>"&objr2("id")&"</td><td>COP VACIO ("&objr("anexo")&")</td>"	
				else
					'response.write "<td>"&objr2("id")&"</td><td>"&objr3("cop")&"</td>"
				end if
				
			end if
			
			
		end if
		response.write "</tr>"
	end if
	objr.movenext
wend
response.write "</table>"



response.end()
'Añadidas, faltan la que le paso en el excel para que les pongan un uso
sql = "select * from temporal_anexo_reach"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table border=1>"
while not objr.eof
	if trim(objr("cas"))<>"" then
		sql2 = "select id from dn_risc_sustancias where num_cas='"&trim(objr("cas"))&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		response.write "<tr>"
		if objr2.eof then
			'No hay ninguno
		else
			'Miro si esa sustantia tiene uso
			sql3="select id, id_uso from dn_risc_sustancias_por_usos where id_sustancia="&objr2("id")
			Set objr3 = Server.CreateObject ("ADODB.Recordset")
			objr3.Open sql3,objconn1,adOpenKeyset
			if not objr3.eof then
				'response.write "<td>"&objr("nombre")&"</td><td>"&objr3("id")&"SI </td><td>"&objr3("id_uso")&"</td>"
			else	
				response.write "<td>"&objr("nombre")&"</td><td>"&objr("cas")&"</td><td></td>"
			end if
			
			
		end if
		response.write "</tr>"
	end if
	objr.movenext
wend
response.write "</table>"
response.end()
'USOS
sql = "select * from temporal_usos order by codigo_uso asc"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table>"
while not objr.eof
	if trim(objr("cas"))<>"" then
		response.write "<tr>"
		response.write "<td>"&objr("uso")&"</td>"	
		'Compruebo si ese numero CAS está en la BBDD
		sql2 = "select id from dn_risc_sustancias where num_cas='"&trim(objr("cas"))&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		if not objr2.eof then
			response.write "<td>"&objr("cas")&"</td>"
			response.write "<td>SI</td>"'Compruebo si ya está tb en la dn_risc_sustancias_por_usos
			'Ahora compruebo si ya está en la lista de usos
			sql3="select id from dn_risc_sustancias_por_usos where id_uso="&trim(objr("codigo_uso"))&" and id_sustancia="&objr2("id")
			Set objr3 = Server.CreateObject ("ADODB.Recordset")
			objr3.Open sql3,objconn1,adOpenKeyset
			if not objr3.eof then
				response.write "<td>ASIGNADA</td>"
			else
				sql_i = "insert into dn_risc_sustancias_por_usos(id_sustancia,id_uso,toxico) values('"&objr2("id")&"','"&objr("codigo_uso")&"',1)"
				response.write "<td>"&sql_i&"</td>"
				
				objconn1.execute(sql_i)
				'response.write "<td>"&sql_i&"</td>"
			end if
			
			
		else
			'Meto esa sustancia en la BBDD
			'sql_i = "insert into temporal_usos_nuevos values('"&unQuote(trim(objr("codigo_uso")))&"','"&unQuote(trim(objr("uso")))&"','"&unQuote(trim(objr("cas")))&"','"&unQuote(trim(objr("nombre")))&"')"
			'objconn1.execute(sql_i)
			'response.write "<tr>"
			'response.write "<td>"&objr("cas")&"</td>"
			'response.write "<td>NO</td>"
			'response.write "<td>"&objr("nombre")&"</td>"
			'response.write "<td>"&sql_i&"</td>"
			'response.write "</tr>"	
		end if

		response.write "</tr>"
			
				
	end if
	objr.movenext
wend
response.write "</table>"


response.end()
'PESTICIDAS

sql = "select * from temporal_pesticidas"

Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table>"
while not objr.eof
	if trim(objr("cas"))<>"" then
		response.write "<tr>"	
		'Compruebo si ese numero CAS está en la BBDD
		sql2 = "select id from dn_risc_sustancias where num_cas='"&trim(objr("cas"))&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		if not objr2.eof then
			response.write "<td>"&objr("cas")&"</td>"
			response.write "<td>SI</td>"'Compruebo si ya está tb en la dn_risc_sustancias_ambiente
			'Ahora compruebo si ya está en la lista de usos
			sql3="select id from dn_risc_sustancias_por_usos where id_uso=2054 and id_sustancia="&objr2("id")
			Set objr3 = Server.CreateObject ("ADODB.Recordset")
			objr3.Open sql3,objconn1,adOpenKeyset
			if not objr3.eof then
				response.write "<td>YA ESTÁ</td>"
			else
				sql_i = "insert into dn_risc_sustancias_por_usos(id_sustancia,id_uso,toxico) values('"&objr2("id")&"','2054',1)"
				objconn1.execute(sql_i)
				response.write "<td>"&sql_i&"</td>"
			end if
			response.write "<td></td>"
			
		else
			'Meto esa sustancia en la BBDD
			'sql_i = "insert into temporal_pesticidas_nuevos values('"&trim(objr("cas"))&"','"&unQuote(trim(objr("nombre")))&"')"
			'objconn1.execute(sql_i)
			'response.write "<td>"&objr("cas")&"</td>"
			'response.write "<td>NO</td>"
			'response.write "<td>"&objr("nombre")&"</td>"

		end if
		
			response.write "</tr>"	
				
		end if
	objr.movenext
wend
response.write "</table>"



response.End()
'PTR (CONTAMINANTES DE SUELO)
sql = "select * from temporal_prtr"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table>"
while not objr.eof
	if trim(objr("cas"))<>"" and trim(objr("suelo")="X") then
		response.write "<tr>"	
		'Compruebo si ese numero CAS está en la BBDD
		sql2 = "select id from dn_risc_sustancias where num_cas='"&trim(objr("cas"))&"'"
		Set objr2 = Server.CreateObject ("ADODB.Recordset")
		objr2.Open sql2,objconn1,adOpenKeyset
		if not objr2.eof then
			response.write "<td>"&objr("cas")&"</td>"
			response.write "<td>SI</td>"'Compruebo si ya está tb en la dn_risc_sustancias_ambiente
			sql3 = "select count(id_sustancia) as conta from dn_risc_sustancias_ambiente where id_sustancia="&objr2("id")
			Set objr3 = Server.CreateObject ("ADODB.Recordset")
			objr3.Open sql3,objconn1,adOpenKeyset
			if objr3("conta")>0 then
				response.write "<td>SI</td>"
			else
				response.write "<td>NO</td>"
			end if
			
		
			sql = "update dn_risc_sustancias_ambiente set eper_suelo=1 where id_sustancia="&objr2("id")
		    objconn1.execute(sql)
			response.write "<td>"&sql&"</td>"
			
		else
	
			response.write "<td>"&objr("cas")&"</td>"
			response.write "<td>NO</td>"
			response.write "<td>"&objr("sustancia")&"</td>"
	
		end if
		
			response.write "</tr>"	
				
		end if
	objr.movenext
wend
response.write "</table>"



response.End()
'CONTAMINANTES PARA EL SUELO
sql = "select * from temporal_contaminantes_suelos"
Set objr = Server.CreateObject ("ADODB.Recordset")
objr.Open sql,objconn1,adOpenKeyset
response.write "<table>"
while not objr.eof
		response.write "<tr>"	
	
	'Compruebo si ese numero CAS está en la BBDD
	sql2 = "select id from dn_risc_sustancias where num_cas='"&trim(objr("cas"))&"'"
	Set objr2 = Server.CreateObject ("ADODB.Recordset")
	objr2.Open sql2,objconn1,adOpenKeyset
	if not objr2.eof then
		response.write "<td>"&objr("cas")&"</td>"
		response.write "<td>SI</td>"
		sql = "update dn_risc_sustancias_ambiente set toxicidad_suelo=1 where id_sustancia="&objr2("id")
		objconn1.execute(sql)
		response.write "<td>"&sql&"</td>"
		
	else

		response.write "<td>"&objr("cas")&"</td>"
		response.write "<td>NO</td>"
		response.write "<td>"&objr("nombre")&"</td>"

	end if
	
		response.write "</tr>"	
			
	
	objr.movenext
wend
response.write "</table>"


%>
