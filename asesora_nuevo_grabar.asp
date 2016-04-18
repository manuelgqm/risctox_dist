<html>
<body>

<!--#include file="unquote.asp"-->

<%

Server.ScriptTimeout = 100000
'
'----- Si es restringida y no estás identificado no puedes entrar
'if session("id_ecogente")="" then response.redirect "acceso.asp?idpagina="&idpagina
'---- ATENCIÓN: ponerlo cuando publiquemos en abierto

Set Upload = Server.CreateObject("Persits.Upload")
Upload.ProgressID = Request.QueryString("PID")

Upload.OverwriteFiles = False
Upload.SetMaxSize 300000000


'ruta_upload_fis = "d:\xvrt\istas.net\html\ecoinformas\ficheros\"
ruta_upload_fis     = "d:\xvrt\istas.net\html\Recursos\"

On Error Resume Next

Count = Upload.Save 

asunto   = Upload.Form("asunto")
pregunta = Upload.Form("pregunta")

'On Error Goto 0

Set File = Upload.Files("myFile")
ext = File.Ext  ' Extensión del fichero

Const adOpenKeyset = 1
DIM objConnection	
DIM objRecordset

Set OBJConnection = Server.CreateObject("ADODB.Connection")
OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

  orden = "INSERT INTO ECOINFORMAS_CONSULTAS (asunto,texto,fecha,usuario,estado,tipo_consulta,fichero) VALUES ("
  orden = orden & "'" & unquote(asunto)        & "',"
  orden = orden & "'" & unquote(pregunta)      & "',"
  orden = orden & "'" & now()                  & "',"
  orden = orden & "'" & session("id_ecogente") & "',"
  orden = orden & "'" & "151"                  & "',"    ' Estado "SIN ASIGNAR"
  orden = orden & "'" & "157"                  & "',"    ' Tipo   "PREGUNTA" 
  orden = orden & "'" & ext                    & "')"    ' Extensión del fichero
  set objRecordset = OBJConnection.Execute(orden)


if count=1 then
  orden = "SELECT max(idconsulta) AS max_id FROM ECOINFORMAS_CONSULTAS "
  set objRecordset = OBJConnection.Execute(orden)
  max_id = objRecordset("max_id")
 

  For Each File in Upload.Files
   File.SaveAs  ruta_upload_fis & "ASESORA_" & max_id & ext
  next 
end if


'
'
'------------------------------------------------------------------------
'
' Enviar email al usuario y al coordinador
'
'
'
' Al usuario(a) que acaba de grabar la pregunta
'idgente = session("id_ecogente")
'modelo  = "0"
'enviado = enviar_mail()

'
'------------------------------------------------------------------------
'
' Al coordinador
'
orden = "SELECT email FROM ECOINFORMAS_GENTE WHERE asesor=2 "
set objRecordset = OBJConnection.Execute(orden)
'
destinatarios = ""
conec = ""
do while not objRecordset.eof
 destinatarios = destinatarios & conec & trim(objRecordset("email"))
 conec = ","
 objRecordset.movenext
loop

modelo = "1"
enviado = enviar_mail()
'
'
'------------------------------------------------------------------------

%>

<!--#include file="asesora_enviar_email.asp"-->

<script>
 //alert ("Enviado:<%=enviado%>");
 opener.location.href='asesora_paso1.asp';
 window.close();
</script>

</body>
</html>

