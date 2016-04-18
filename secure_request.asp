<%
function clearSQL(strTexto)
	Dim cadenas_chungas, strTexto2, ataque, i
	
	strTexto = Replace(strTexto,"'","''")
	cadenas_chungas = array("--","script","select","drop","update","insert","delete","xp_","exec","cast(","sys","sp_")    
	ataque = ""
	strTexto2 = strTexto
	
	for i = 0 to uBound(cadenas_chungas)    
		if inStr(lcase(strTexto),lcase(cadenas_chungas(i)))>0 then ataque = "SQL injected"
		'if len(strTexto)>100 and inStr(request.ServerVariables("URL"),"articulo.asp")=0 then ataque="Too long string"
		strTexto2 = replace(strTexto2, cadenas_chungas(i), "",1,-1,1)    
	next    
	
	clearSQL = strTexto2
	
	if ataque<>"" then
		call notifyAttack( strTexto, ataque )
	end if
	
end function

function requestClean( strTexto )
	requestClean = clearSQL( request( strTexto ) )
end function

function executeSecure( strScript )
	
	if instr( strScript, ":" ) or len( strScript ) > 30 then
		call notifyAttack( strScript, "Code injection" )
	else
		execute( strScript )
	end if
	
end function

sub notifyAttack( strTexto, kind )
	Dim cadena0, cadena1, cadena2, cadena3, IP, Mail
	cadena0 = request.ServerVariables("HTTP_HOST")
	cadena1 = request.ServerVariables("URL")
	cadena2 = request.ServerVariables("QUERY_STRING")
	cadena3 = request.ServerVariables("ALL_HTTP")
	IP = Request.ServerVariables("REMOTE_ADDR")
	Set Mail = Server.CreateObject("Persits.MailSender")
	Mail.Host = "localhost"
	Mail.From = "soporteapps@istas.net"
	Mail.FromName = Mail.EncodeHeader("SERVIDOR") ' Opcional 
	Mail.AddAddress "soporteapps@istas.net"
	Mail.Subject = Mail.EncodeHeader("Ataque: " & now() & " tipo: " & kind &" IP: " & IP)
	Mail.Body = "http://" & cadena0 & cadena1 & "?" & cadena2 & "<br />" & IP & "<br />CADENA=" & strTexto & "<br />" & replace(cadena3,chr(10),"<br />")
	Mail.IsHTML = True
	On Error Resume Next
	Mail.SendToQueue		
end sub
%>