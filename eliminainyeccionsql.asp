<%

FUNCTION EliminaInyeccionSQL(strTexto)
'response.write "<br>strTexto: " & strTexto
	
	strTexto = Replace(strTexto,"'","''")
	cadenas_chungas = array("--","script","select","drop","update","insert","delete","xp_","exec","cast(","sys","sp_")    
	ataque = ""
	strTexto2 = strTexto
	for cont = 0 to uBound(cadenas_chungas)    
		if inStr(lcase(strTexto),lcase(cadenas_chungas(cont)))>0 then ataque="cadena chunga"
		if len(strTexto)>100 and request.ServerVariables("URL")<>"/pe/articulo.asp" then ataque="cadena larga"
		strTexto2 = replace(strTexto2, cadenas_chungas(cont), "",1,-1,1)
	next    
	
	EliminaInyeccionSQL = strTexto2
	if 1=0 and ataque<>"" then
'response.write "<br>ataque: " & ataque
		cadena = request.ServerVariables("QUERY_STRING")
		cadena2 = request.ServerVariables("URL")
		IP = Request.ServerVariables("REMOTE_ADDR")
		cadena3 = request.ServerVariables("ALL_HTTP")
		Set Mail = Server.CreateObject("Persits.MailSender")
		Mail.Host = "localhost"
		Mail.From = "xip@xipmultimedia.com"
		Mail.FromName = Mail.EncodeHeader("SERVIDOR") ' Opcional 
		Mail.AddAddress "xavi@xipmultimedia.com"
		Mail.Subject = Mail.EncodeHeader("Ataque: "&now()&" tipo: "&ataque&" IP: "&IP)
		Mail.Body = "www.istas.net"&cadena2&"?"&cadena&"<br />"&IP&"<br />CADENA="&strTexto&"<br />"&replace(cadena3,chr(10),"<br />")
		Mail.IsHTML = True
		On Error Resume Next
		Mail.SendToQueue		
	end if
	
END FUNCTION 

%>