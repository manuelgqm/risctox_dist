<%

	Server.ScriptTimeout = 100000
	
	set objConnection2 = Server.CreateObject("ADODB.Connection")
	objConnection2.open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("atp30_31.mdb")
	
	set objConnection = Server.CreateObject("ADODB.Connection")
	OBJConnection.connectionstring="Provider=SQLOLEDB; Data Source=DISOLTEC03\XIP; Initial Catalog=istas_risctox081009; User ID=istas_sql_usuari; Password=***REMOVED**"
	OBJConnection.Open

	sql = "select top 700 * from atp30_31 where id>700 order by id"
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
		
		for sus=0 to num_sustancias-1
		sustancias = sustancias+1
		
		if num_sustancias>1 then 
			cas = trim(cas_(sus))
			ce = trim(ce_(sus))
			nombre = trim(nombre_(sus))
			nombre_ing = trim(nombre_ing_(sus))
		end if
		cas = trim(cas)
		ce = trim(ce)
		cas = replace(cas,chr(10),"")
		ce = replace(ce,chr(10),"")
		cas = replace(cas,chr(32),"")
		ce = replace(ce,chr(32),"")
		sql2 = "select id,num_cas,num_ce_einecs from dn_risc_sustancias where (num_cas='"&cas&"' and num_cas<>'') or (num_ce_einecs='"&ce&"' and num_ce_einecs<>'')"
		Set DSql2 = OBJConnection.Execute(sql2)
		if not dsql2.eof then 
			id_sustancia_risctox = dsql2("id")
			num_cas = dsql2("num_cas")
			num_ce_einecs = dsql2("num_ce_einecs")
			
			if (num_cas<>"" and cas="") or (num_ce_einecs<>"" and ce="") then
				sql3 = "update dn_risc_sustancias set "
				if (num_cas<>"" and cas="") then sql3 = sql3 & "num_cas='"&num_cas&"' "
				if (num_ce_einecs<>"" and ce="") then sql3 = sql3 & "num_ce_einecs='"&num_ce_einecs&"' "
				sql3 = sql3 & " where id="&id_sustancia_risctox
				response.write sql3 & "<br />"
				'Set DSql3 = OBJConnection.Execute(sql3)
			end if
				
		end if
	next
	dsql.movenext
	loop
	
	
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