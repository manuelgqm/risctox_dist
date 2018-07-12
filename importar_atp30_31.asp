<%

	Server.ScriptTimeout = 100000
	
	set objConnection2 = Server.CreateObject("ADODB.Connection")
	objConnection2.open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("atp30_31.mdb")
	
	set objConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox; User ID=istas_sql_usuari; Password=***REMOVED**"
	OBJConnection.Open
	salto = chr(10)&chr(13)
	coinciden = 0
	dim clasif_nuevas()
	redim clasif_nuevas(15)
	dim concent_nuevas(),eti_nuevas()
	redim concent_nuevas(15)
	redim eti_nuevas(15)		
	response.write "<b>Importación a RISCTOX con fecha "&date()&"</b><br />"	
	response.write "Leyenda: ATP 30 y 31, <span style='color:#ff0000'>lo que hay en la RISCTOX con el mismo CAS o CE</span> y <span style='color:#0000ff'>lo que va a haber</span><br /><br />"
	response.write "<table style='font-style:courier;font-size:9pt' border=1><tr><td>Núm</td><td>Núm anterior</td><td>Nombre</td><td>Sinónimos</td><td>Nombre_ing</td><td>CAS</td><td>CE</td><td>RD</td>"
	response.write "<td>Símbolos</td><td>Frases S</td>"
	for i=1 to 15
		response.write "<td nowrap>Clasificación "&i&"</td>"
	next
	for i=1 to 15
		response.write "<td nowrap>Concretación "&i&"</td><td nowrap>Etiquetado "&i&"</td>"
	next
	response.write "<td nowrap>Notas RD 363</td></tr>"
	
	'sustancias = 763
	'sql = "select top 1000 * from atp30_31 where id>700 order by id"
	
	sustancias = 0
	sql = "select top 300 * from atp30_31 where id>1400 order by id"
	Set DSql = Server.CreateObject("ADODB.Recordset")
	Set DSql = OBJConnection2.Execute(sql)
	do while not dsql.eof
		cas = sin_corchetes(dsql("cas"))
		ce = sin_corchetes(dsql("ec"))
		nombre = sin_corchetes(dsql("nombre"))
		nombre_ing = sin_corchetes(dsql("nombre_ing"))
		nombre_ing = replace(nombre_ing,".sb.","")
		if instr(cas,"xavi")<>0 then
			cas_ = split(cas,"xavi")
			ce_ = split(ce,"xavi")
			nombre_ = split(nombre,"xavi")
			nombre_ing_ = split(nombre_ing,"xavi")
			num_sustancias = ubound(cas_)
		else
			num_sustancias = 1
		end if
		indice = trim(dsql("indice"))
		etiquetado = arregla(dsql("etiquetado"))
		etiq = split(etiquetado,"<br />")
		num_etiq = ubound(etiq)
		if num_etiq>0 then simbolos = etiq(0)
		if num_etiq>=1 then frases_r = etiq(1)
		if num_etiq>=2 then frases_s = etiq(2)
		frases_s = replace(frases_s,"S: ","")
		clasificacion = arregla(dsql("clasificacion"))
		clasif = split(clasificacion,"<br />")
		num_clasif = ubound(clasif)
		concentracion = arregla(dsql("concentracion"))
		concent = split(concentracion,"<br />")
		num_concent = ubound(concent)
		notas1 = trim(dsql("notas1"))
		if isnull(notas1) then notas1=""
		notas1 = replace(notas1," ","")
		notas1 = replace(notas1,".","")
		notas1 = replace(notas1,chr(10),"")
		notas1_ = ""
		for n=1 to len(notas1)
			notas1_ = notas1_ & "Nota "&mid(notas1,n,1)&". "
		next
		notas2 = trim(dsql("notas2"))
		if isnull(notas2) then notas2=""
		notas2 = replace(notas2," ","")
		notas2 = replace(notas2,".","")
		notas2 = replace(notas2,chr(10),"")
		notas2_ = ""
		for n=1 to len(notas2)
			notas2_ = notas2_ & "Nota "&mid(notas2,n,1)&". "
		next		
		notas = notas1_ & notas2_
		
		for sus=0 to num_sustancias-1
		sustancias = sustancias+1
		
		if num_sustancias>1 then 
			cas = trim(cas_(sus))
			ce = trim(ce_(sus))
			nombre = trim(nombre_(sus))
			nombre_ing = trim(nombre_ing_(sus))
		end if
			
		response.write "<tr><td valign=top>"&sustancias& "</td><td valign=top>"&dsql("id")& "</td><td valign=top>"&nombre&"</td><td></td><td valign=top>"&nombre_ing&"</td><td valign=top nowrap>"&cas&"</td><td valign=top nowrap>"&ce&"</td><td valign=top>"&indice&"</td>"
		response.write "<td valign=top>"&simbolos&"</td>"
		response.write "<td valign=top>"&frases_s&"</td>"
			
		cl = 0
		i = 0

		do while i<15 and cl<15
			if i<=num_clasif then 
				clasif_ = clasif(i)
				if instr(clasif_,"-")<>0 and instr(clasif_,"R50-53")=0  and instr(clasif_,"R51-53")=0 and instr(clasif_,"R52-53")=0 then
					clasif_v = split(clasif_,";")
					clasif_0 = clasif_v(0)
					clasif_1 = clasif_v(1)
					clasif_1_v = split(clasif_1,"-")
					response.write "<td>"&clasif_0&"; "&clasif_1_v(0)&"</td><td>"&clasif_0&"; R"&clasif_1_v(1)&"</td>"
					clasif_nuevas(cl) = clasif_0&"; "&clasif_1_v(0)
					clasif_nuevas(cl+1) = clasif_0&"; R"&clasif_1_v(1)
					cl = cl+2
				else
					response.write "<td>"&clasif(i)&"</td>"
					clasif_nuevas(cl) = clasif(i)
					cl = cl+1
				end if
			else
				response.write "<td> </td>"
				clasif_nuevas(cl) = ""
				cl = cl+1
			end if
			i=i+1			
		loop
		
		for i=0 to 14
			if i<num_concent then
				linea_concent = concent(i)
				linea_concent1 = split(linea_concent,":")
				concent_ = linea_concent1(0)
				eti_con_ = linea_concent1(1)
				response.write "<td>"&concent_&"</td><td>"&eti_con_&"</td>"
				concent_nuevas(i) = concent_
				eti_nuevas(i) = eti_con_
			else
				response.write "<td> </td><td> </td>"
				concent_nuevas(i) = ""
				eti_nuevas(i) = ""
			end if
		next
		
		response.write "<td nowrap>"&notas&"</td></tr>"&salto

		'--lo que hay en la risctox
		nombre_grabado = ""
		nombre_ing_grabado = ""
		cas = trim(cas)
		ce = trim(ce)
		cas = replace(cas,chr(10),"")
		ce = replace(ce,chr(10),"")
		cas = replace(cas,chr(32),"")
		ce = replace(ce,chr(32),"")
		sql2 = "select * from dn_risc_sustancias where (num_cas='"&cas&"' and num_cas<>'') or (num_ce_einecs='"&ce&"' and num_ce_einecs<>'')"
		'response.write "<tr><td colspan=200>cas"&len(cas)&"="
		'for a=1 to len(cas)
		'	response.write asc(mid(cas,a,1))&"-"
		'next	
		'response.write "</td></tr>"
		Set DSql2 = OBJConnection.Execute(sql2)
		
		if not dsql2.eof then 
			esta_en_risctox = "si"
			coinciden = coinciden+1
			nombre_grabado = dsql2("nombre")
			nombre_ing_grabado = dsql2("nombre_ing")
			id_sustancia_risctox = dsql2("id")
			response.write "<tr><td></td><td></td><td valign=top style='color:#ff0000'>"&dsql2("nombre")& "</td><td></td>"
			response.write "<td valign=top style='color:#ff0000'>"&dsql2("nombre_ing")&"</td>"
			response.write "<td valign=top style='color:#ff0000'>"&dsql2("num_cas")&"</td>"
			response.write "<td valign=top style='color:#ff0000'>"&dsql2("num_ce_einecs")&"</td>"
			response.write "<td valign=top style='color:#ff0000'>"&dsql2("num_rd")&"</td>"
			response.write "<td valign=top style='color:#ff0000'>"&dsql2("simbolos")&"</td>"
			response.write "<td valign=top style='color:#ff0000'>"&dsql2("frases_s")&"</td>"
			for i=1 to 15
				response.write "<td valign=top style='color:#ff0000'>"&dsql2("clasificacion_"&i)&"</td>"
			next
			for i=1 to 15
				response.write "<td valign=top style='color:#ff0000'>"&dsql2("conc_"&i)&"</td><td valign=top style='color:#ff0000'>"&dsql2("eti_conc_"&i)&"</td>"
			next
			response.write "<td valign=top style='color:#ff0000'>"&dsql2("notas_rd_363")&"</td></tr>"&salto
		else
			esta_en_risctox = "no"
		end if
		
		'--lo que se graba
		response.write "<tr><td></td><td></td><td valign=top style='color:#0000ff'>"&nombre& "</td>"
		response.write "<td valign=top style='color:#0000ff'>"
		if lcase(nombre)<>lcase(nombre_grabado) and instr(lcase(nombre),lcase(nombre_grabado))=0 then 
			sinonimo = nombre_grabado
			response.write nombre_grabado
		else
			sinonimo = ""
		end if
		response.write "</td>"
		response.write "<td valign=top style='color:#0000ff'>"
		if nombre_ing_grabado="" then 
			nombre_ing_nuevo = nombre_ing
			response.write nombre_ing 
		else 
			if trim(lcase(nombre_ing_grabado))<>trim(lcase(nombre_ing)) then 
				nombre_ing_nuevo = nombre_ing&"@"&nombre_ing_grabado
				response.write nombre_ing&"@"&nombre_ing_grabado
			else
				nombre_ing_nuevo = nombre_ing
				response.write nombre_ing
			end if
		end if
		response.write "</td>"
		response.write "<td valign=top style='color:#0000ff'>"&cas&"</td>"
		response.write "<td valign=top style='color:#0000ff'>"&ce&"</td>"
		response.write "<td valign=top style='color:#0000ff'>"&indice&"</td>"
		response.write "<td valign=top style='color:#0000ff'>"&simbolos&"</td>"
		response.write "<td valign=top style='color:#0000ff'>"&frases_s&"</td>"
		for i=0 to 14
			response.write "<td valign=top style='color:#0000ff'>"&clasif_nuevas(i)&"</td>"
		next
		for i=0 to 14
			response.write "<td valign=top style='color:#0000ff'>"&concent_nuevas(i)&"</td><td valign=top style='color:#0000ff'>"&eti_nuevas(i)&"</td>"
		next
		response.write "<td valign=top style='color:#0000ff'>"&notas&"</td></tr>"&salto

		if esta_en_risctox="si" then 
			sql3 = "update dn_risc_sustancias set nombre='"&unquote(left(nombre,1500))&"',nombre_ing='"&unquote(left(nombre_ing_nuevo,1500))&"',num_cas='"&cas&"',num_ce_einecs='"&ce&"',num_rd='"&indice&"',simbolos='"&simbolos&"',frases_s='"&frases_s&"'"
			for i=0 to 14
				sql3 = sql3 & ",clasificacion_"&cstr(i+1)&"='"&unquote(clasif_nuevas(i))&"'"
			next
			for i=0 to 14
				sql3 = sql3 & ",conc_"&cstr(i+1)&"='"&unquote(concent_nuevas(i))&"',eti_conc_"&cstr(i+1)&"='"&unquote(eti_nuevas(i))&"'"
			next
			sql3 = sql3 & ",notas_rd_363='"&unquote(notas)&"'" 
			sql3 = sql3 & " where id="&id_sustancia_risctox
		end if

		if esta_en_risctox="no" then 
			sql3 = "insert into dn_risc_sustancias (nombre,nombre_ing,num_cas,num_ce_einecs,num_rd,simbolos,frases_s,"
			sql3 = sql3 & "clasificacion_1,clasificacion_2,clasificacion_3,clasificacion_4,clasificacion_5,clasificacion_6,clasificacion_7,clasificacion_8,clasificacion_9,clasificacion_10,clasificacion_11,clasificacion_12,clasificacion_13,clasificacion_14,clasificacion_15,"
			sql3 = sql3 & "conc_1,eti_conc_1,conc_2,eti_conc_2,conc_3,eti_conc_3,conc_4,eti_conc_4,conc_5,eti_conc_5,conc_6,eti_conc_6,conc_7,eti_conc_7,conc_8,eti_conc_8,conc_9,eti_conc_9,conc_10,eti_conc_10,conc_11,eti_conc_11,conc_12,eti_conc_12,conc_13,eti_conc_13,conc_14,eti_conc_14,conc_15,eti_conc_15,"
			sql3 = sql3 & "notas_rd_363) values ("
			sql3 = sql3 & "'"&unquote(left(nombre,1500))&"','"&unquote(left(nombre_ing_nuevo,1500))&"','"&unquote(cas)&"','"&unquote(ce)&"','"&unquote(indice)&"','"&unquote(simbolos)&"','"&unquote(frases_s)&"','"
			for i=0 to 14
				sql3 = sql3 & unquote(clasif_nuevas(i))&"','"
			next
			for i=0 to 14
				sql3 = sql3 & unquote(concent_nuevas(i))&"','"&unquote(eti_nuevas(i))&"','"
			next
			sql3 = sql3 & unquote(notas) & "')"
		end if
		
		response.write "<tr><td colspan=100>"&sql3&"</td></tr>"
		Set DSql3 = OBJConnection.Execute(sql3)
		
		if sinonimo<>"" then
			sql4 = "select id from dn_risc_sinonimos where id_sustancia="&id_sustancia_risctox&" and nombre='"&unquote(sinonimo)&"'"
			Set DSql4 = OBJConnection.Execute(sql4)
			if not dsql4.eof then
				sql5 = "insert into dn_risc_sinonimos (id_sustancia,nombre) values ('"&id_sustancia_risctox&"','"&unquote(sinonimo)&"');"
				Set DSql5 = OBJConnection.Execute(sql5)
			end if
		end if
			
		next
		dsql.movenext
	loop
	response.write "</table>"
	response.write "<br />Sustancias ATP 30 y 31: "&sustancias
	response.write "<br />Coinciden (se actualizan): "&coinciden
	response.write "<br />Nuevas (se introducen): "&cstr(sustancias-coinciden)
	
	function arregla(algo_)
		if isnull(algo_) then 
			algo = ""
		else
			algo = algo_
		end if
		algo = trim(algo)
		if isnull(algo) then algo=""
		algo = replace(algo,chr(10),"<br />")
		arregla = algo
	end function
	
	
	function sin_corchetes(algo_)
		algo = trim(algo_)
		algo = replace(algo,"[1]","xavi")
		algo = replace(algo,"[2]","xavi")
		algo = replace(algo,"[3]","xavi")
		algo = replace(algo,"[4]","xavi")
		algo = replace(algo,"[5]","xavi")
		algo = replace(algo,"[6]","xavi")
		algo = replace(algo,"[7]","xavi")
		algo = replace(algo,"[8]","xavi")
		algo = replace(algo,"[9]","xavi")
		algo = replace(algo,"[10]","xavi")
		algo = replace(algo,"[11]","xavi")
		algo = replace(algo,"[12]","xavi")
		algo = replace(algo,"[13]","xavi")
		algo = replace(algo,"[14]","xavi")
		algo = replace(algo,"[15]","xavi")
		algo = replace(algo,"[16]","xavi")
		algo = replace(algo,"[17]","xavi")
		algo = replace(algo,"[18]","xavi")
		algo = replace(algo,"[19]","xavi")
		algo = replace(algo,"[20]","xavi")
		sin_corchetes = algo
	end function
	
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

%>