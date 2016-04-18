<%
	tipo = clng(Request("tipo"))
	iden = session("id_ecogente")
	id = clng(Request("id"))
	formasunto = unquote(Request.Form("asunto"))
	formtexto = unquote(Request.Form("texto"))
	if not (cstr(formtexto)="" and cstr(formasunto)="") then
		Set OBJConnection = Server.CreateObject("ADODB.Connection")
		OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

		'--- 1º Recuperar el id afectado y guardarse el siguiente y nivel...
		orden = "SELECT sig,nivel FROM ECO_FOROS where id="&id
		Set Dorga = OBJConnection.Execute(orden)
		
		sig_cambiar = clng(dorga("sig"))
		niv_cambiar = clng(dorga("nivel"))+1

		'--- 2º Insertar nuevo mensaje, poner el puntero siguiente el que tenía su padre...
		orden = "INSERT into ECO_FOROS (nivel,sig,asunto,idgente,fecha,texto,tipo,adjunto) " & _
	        " values (" & niv_cambiar & ", " & sig_cambiar & ",'" & formasunto & "'," & iden & ", '" & now() & "', '" & formtexto & "'," & tipo & ",'');"
	        Set Dorga = OBJConnection.Execute(orden)
		
		orden = "SELECT max(id) AS nuevo_id FROM ECO_FOROS WHERE tipo="&tipo
		Set Dorga = OBJConnection.Execute(orden)
		id_nuevo=clng(Dorga("nuevo_id"))	        
		
		'--- 3º Actualizar el padre poniendo como siguiente el actual registro insertado...
		orden = "UPDATE ECO_FOROS SET sig="&id_nuevo&" WHERE id="&id
		Set Dorga = OBJConnection.Execute(orden)
		
		'--- 4º Marcar como leído
		orden = "INSERT into ECO_FOROS_LEIDOS (idmensaje,idgente,fecha) VALUES ("&id_nuevo&","&iden&",'"&now()&"')"
		Set Dorga = OBJConnection.Execute(orden)
				
	end if

'-------------------------------
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

<html>

<head><title>Publicar anuncio</title></head>


<body bgcolor="#FFFFFF" onload="javascript:vamos()">
ok, mensaje publicado
</body>
</html>

<script LANGUAGE="JScript">

function vamos() {

window.opener.location.reload();
window.close();

}

</script>
