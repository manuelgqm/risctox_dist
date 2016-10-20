<%
function sanitizeScript( strScript )
	if instr( strScript, ":" ) or len( strScript ) > 30 then
		call notifyAttack( strScript, "Suspected code injection" )
		sanitizeExecution = ""
		exit function
	end if
	
	sanitizeScript = strScript
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