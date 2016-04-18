
<html>
<body>

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

ruta_upload_fis     = "d:\xvrt\istas.net\html\Recursos\"

On Error Resume Next

Count = Upload.Save 

asesor              = Upload.Form("asesor")

asunto              = Upload.Form("asunto")
pregunta            = Upload.Form("pregunta")
idconsulta          = Upload.Form("idconsulta")

'response.redirect "comprueba.asp?orden="&asesor

asesor              = Upload.Form("asesor")      ' Tipo de persona
estado_consulta_pri = Upload.Form("estado_consulta_pri")
tema_consulta_fil   = Upload.Form("tema_consulta_fil")
act_pag             = Upload.Form("act_pag")
estado_consulta     = Upload.Form("estado_consulta")
tema_consulta       = Upload.Form("tema_consulta")

' Grabar la respuesta y el fichero adjunto

if count=1 then
 Set File = Upload.Files("myFile")
 ext = File.Ext  ' Extensión del fichero
end if 
'

Const adOpenKeyset = 1
DIM objConnection	
DIM objRecordset

Set OBJConnection = Server.CreateObject("ADODB.Connection")
OBJConnection.Open "DSN=istas.net.base;UID=qc507;PWD=sql"

if pregunta<>"" then
   orden = "INSERT INTO ECOINFORMAS_CONSULTAS (asunto,texto,fecha,usuario,tipo_consulta,puntero,fichero) VALUES ("
   orden = orden & "'" & unquote(asunto)        & "',"
   orden = orden & "'" & unquote(pregunta)      & "',"
   orden = orden & "'" & now()                  & "',"
   orden = orden & "'" & session("id_ecogente") & "',"
   orden = orden & "'" & "158"                  & "',"    ' Tipo consulta: RESPUESTA
   orden = orden & "'" & idconsulta             & "',"    ' Puntero
   orden = orden & "'" & ext                    & "')"    ' Extensión del fichero
   set objRecordset = OBJConnection.Execute(orden)

   orden = "SELECT max(idconsulta) AS max_id FROM ECOINFORMAS_CONSULTAS "
   set objRecordset = OBJConnection.Execute(orden)
   max_id = objRecordset("max_id")
end if

if idconsulta<>"" then
 orden = "UPDATE ECOINFORMAS_CONSULTAS SET estado='"&estado_consulta&"',tema_consulta='"&tema_consulta&"' WHERE idconsulta="&idconsulta
 set objRecordset = OBJConnection.Execute(orden)  
end if 

if count=1 and max_id<>"" then
  For Each File in Upload.Files
   File.SaveAs  ruta_upload_fis & "ASESORA_" & max_id & ext
  next 
end if

' ------------------------------------------------------------------------------------------------------------------------------------------
'
' Es el coordinador el que hace la asignación de tema.
' Habrá que crear un registro en la ficha SAAT con los datos que podamos sacar
'
if asesor="2" and idconsulta<>"" then

 '---------------------------------------
 ' Asignaré todas las variables de la tabla ECOINFORMAS_GENTE para que sea más fácil la asignación al resto de tablas de SAAT
 '---------------------------------------  
 
 ' Asignar los datos de la consulta: idusuario, asunto y consulta
 orden = "Select * from ECOINFORMAS_CONSULTAS where idconsulta="&idconsulta
 set objRecordset = OBJConnection.Execute(orden) 
 
 eco_con_idusuario   = objRecordset("usuario")
 eco_con_fecha       = objRecordset("fecha") 
 eco_con_asunto      = unquote(objRecordset("asunto"))
 eco_con_consulta    = unquote(objRecordset("texto"))
 
 ' Voy a juntar el asunto con la consulta
 eco_con_consulta    = left((eco_con_asunto&chr(10)&chr(13)&eco_con_consulta),3000)

 ' Asignar los datos de empresa y del trabajador
 orden = "Select * from ECOINFORMAS_GENTE where idgente='"&eco_con_idusuario&"'"
 Set objRecordset = Server.CreateObject ("ADODB.Recordset")
 objRecordset.Open orden,objConnection,adOpenKeyset

 eco_tra_nombre      = objRecordset("nombre")
 eco_tra_apellidos   = unquote(left(objRecordset("apellidos"),35))
 eco_tra_fec_nac     = objRecordset("fec_nac")  
 eco_tra_sexo        = objRecordset("sexo")     
 eco_tra_direccion   = unquote(left(objRecordset("direccion"),50))
 eco_tra_localidad   = unquote(left(objRecordset("localidad"),50))
 eco_tra_cp          = left(objRecordset("cp"),5)
 eco_tra_telefono    = unquote(left(objRecordset("telefono"),25))
 eco_tra_movil       = unquote(objRecordset("movil"))    
 eco_tra_fax         = unquote(objRecordset("fax"))      
 eco_tra_email       = unquote(left(objRecordset("email"),50))

 xtra_provincia      = objRecordset("provincia")
 eco_tra_provincia   = dame_provincia(xtra_provincia)  ' Como tengo que asignar el valor de la tabla SAAT_VALORES
 eco_tra_territorio  = dame_territorio(eco_tra_provincia)

 orden = "Select * from ECOINFORMAS_GENTE where idgente="&eco_con_idusuario
 set objRecordset2 = OBJConnection.Execute(orden)  

 eco_empresa         = unquote(left(objRecordset2("empresa"),50))
 eco_razon_social    = unquote(left(objRecordset2("razon_social"),50))
 eco_emp_direccion   = unquote(left(objRecordset2("emp_direccion"),50))
 eco_emp_localidad   = unquote(left(objRecordset2("emp_localidad"),50))
 eco_emp_cp          = left(trim(objRecordset2("emp_cp")),5)
 eco_emp_telefono    = unquote(left(objRecordset2("emp_telefono"),25))
 eco_emp_movil       = unquote(objRecordset2("emp_movil"))    
 eco_emp_fax         = unquote(objRecordset2("emp_fax"))      
 eco_emp_email       = unquote(objRecordset2("emp_email"))

 xemp_provincia      = objRecordset2("emp_provincia")
 eco_emp_provincia   = dame_provincia(xemp_provincia)   ' Como tengo que asignar el valor de la tabla SAAT_VALORES
 eco_emp_territorio  = dame_territorio(eco_emp_provincia)  ' Como tengo que asignar el valor de la tabla SAAT_VALORES

 '---------------------------------------
 ' Buscaré el asesor que tiene asignado el tema en Ecoinformas_Valores, debe de estar assignado el asesor.
 ' Falta ver qué pasa si hay más de un asesor para un tema.
 '--------------------------------------- 
 orden = "Select desc2 from ECOINFORMAS_VALORES where valor="&tema_consulta
 set objRecordset = OBJConnection.Execute(orden) 
 if not objRecordset.eof then
  idasesor=objRecordset("desc2")
 end if
 
 if idasesor<>"" then
 
 '--------------------------------------- 
 ' Crear la empresa en SAAT_EMPRESAS   
 '---------------------------------------

 ordensql = "INSERT INTO SAAT_EMPRESAS ("
 ordensql = ordensql & "idasesor,centro,empresa,domicilio,localidad,cp,telefono1,cod_pro,territorio,cod_fed,cod_cna,"
 ordensql = ordensql & "tramo,afi_con,ser_aje,ser_aje2,ser_pro,ser_ajecheck,git_con,procedencia "

 ordensql = ordensql & " ) VALUES ( "

 ordensql = ordensql & "'" & idasesor             & "',"
 ordensql = ordensql & "'" & eco_razon_social     & "',"
 ordensql = ordensql & "'" & eco_empresa          & "',"
 ordensql = ordensql & "'" & eco_emp_direccion    & "',"
 ordensql = ordensql & "'" & eco_emp_localidad    & "',"
 ordensql = ordensql & "'" & eco_emp_cp           & "',"
 ordensql = ordensql & "'" & eco_emp_telefono     & "',"
 ordensql = ordensql & "'" & eco_emp_provincia    & "',"
 ordensql = ordensql & "'" & eco_emp_territorio   & "',"
 ordensql = ordensql & "'" & cod_fed              & "'," ' ?
 ordensql = ordensql & "'" & cod_cna              & "'," ' ?
 ordensql = ordensql & "'" & tramo                & "'," ' ?
 ordensql = ordensql & "'" & afi_con              & "'," ' ?
 ordensql = ordensql & "'" & ser_aje              & "'," ' ?
 ordensql = ordensql & "'" & ser_aje2             & "'," ' ?
 ordensql = ordensql & "'" & ser_pro              & "'," ' ?
 ordensql = ordensql & "'" & ser_ajecheck         & "'," ' ?
 ordensql = ordensql & "'" & git_con              & "'," ' ?
 ordensql = ordensql & "'" & "4"                  & "' " ' Procedencia fija con valor 4 

 ordensql = ordensql & " ) "

 orden_empresa = ordensql

 set objRecordset = OBJConnection.Execute(ordensql) 
 
 orden = "select max(idempresa) as idempm from SAAT_EMPRESAS"
 set objRecordset = OBJConnection.Execute(orden)   
 idemp = objRecordset("idempm")


 '---------------------------------------
 'Crear registro con el trabajador en SAAT_TRABAJADORES
 '--------------------------------------- 

 ordensql = "INSERT INTO SAAT_TRABAJADORES ("
 ordensql = ordensql & "idasesor,nombre,apellidos,domicilio,localidad,cp,telefono1,cod_pro,territorio,email,afiliado,cod_rep,cod_for,procedencia"

 ordensql = ordensql & " ) VALUES ( "

 ordensql = ordensql & "'" & idasesor            & "',"
 ordensql = ordensql & "'" & eco_tra_nombre      & "',"
 ordensql = ordensql & "'" & eco_tra_apellidos   & "',"
 ordensql = ordensql & "'" & eco_tra_direccion   & "',"
 ordensql = ordensql & "'" & eco_tra_localidad   & "',"
 ordensql = ordensql & "'" & eco_tra_cp          & "',"
 ordensql = ordensql & "'" & eco_tra_telefono    & "',"
 ordensql = ordensql & "'" & eco_tra_cod_pro     & "',"
 ordensql = ordensql & "'" & eco_tra_territorio  & "',"
 ordensql = ordensql & "'" & eco_tra_email       & "',"
 ordensql = ordensql & "'" & afiliado            & "'," ' ?
 ordensql = ordensql & "'" & cod_rep             & "'," ' ?
 ordensql = ordensql & "'" & cod_for             & "'," ' ?
 ordensql = ordensql & "'" & "4"                 & "' " ' Procedencia fija con valor 4 

 ordensql = ordensql & " ) "
 
 orden_trabajador = ordensql 
 
 set objRecordset = OBJConnection.Execute(ordensql) 
 
 orden = "select max(idtra) as idtram from SAAT_TRABAJADORES"
 set objRecordset = OBJConnection.Execute(orden)   
 idtra = objRecordset("idtram")
 
 '--------------------------------------- 
 ' Crear registro de la consulta en SAAT_CONSULTAS  
 '--------------------------------------- 

 orden = "SELECT max(numcon) AS numcon FROM SAAT_CONSULTAS WHERE idase="&idasesor
 set objRecordset = OBJConnection.Execute(orden)  
 if not objRecordset.eof then
  if not isnull(objRecordset("numcon")) then
   numcon = clng(objRecordset("numcon"))+1
  else
   numcon = 0
  end if 
 else
  numcon = 0 
 end if 
 
 
  ordensql = "INSERT into SAAT_CONSULTAS ("
  ordensql = ordensql & "idase,consulta,idtra,idemp,fec_ini,cod_for,cod_ase,ori_dem,tiempo,imp_mut,sin_info,otrascausas,"
  ordensql = ordensql & "estado,otc,otr_cambio,res_gestion,not_experiencia,numcon,fec_ent,hor_ent,procedencia"
  ordensql = ordensql & ") VALUES ("
  
  ordensql = ordensql & "'" & idasesor         & "',"
  ordensql = ordensql & "'" & eco_con_consulta & "',"
  ordensql = ordensql & "'" & idtra            & "',"
  ordensql = ordensql & "'" & idemp            & "',"
  ordensql = ordensql & "'" & eco_con_fecha    & "',"
  ordensql = ordensql & "'" & "844"            & "'," ' Forma de consulta: EcoInformas.
  ordensql = ordensql & "'" & tc               & "'," ' ?  
  ordensql = ordensql & "'" & ori_dem          & "'," ' ?  
  ordensql = ordensql & "'" & tiempo           & "'," ' ?  
  ordensql = ordensql & "'" & imp_mut          & "'," ' ?  
  ordensql = ordensql & "'" & sin_info         & "'," ' ?  
  ordensql = ordensql & "'" & otrascausas      & "'," ' ?  
  ordensql = ordensql & "'" & "0"              & "'," ' Abierto
  ordensql = ordensql & "'" & otc              & "'," ' ?
  ordensql = ordensql & "'" & otr_cambio       & "'," ' ?
  ordensql = ordensql & "'" & res_gestion      & "'," ' ?
  ordensql = ordensql & "'" & not_experiencia  & "'," ' ?
  ordensql = ordensql & "'" & numcon           & "'," ' Nº de consulta. Asignado anteriormente
  ordensql = ordensql & "'" & date()	       & "',"
  ordensql = ordensql & "'" & time()	       & "',"
  ordensql = ordensql & "'" & "4"	       & "' " ' Procedencia fija con valor 4 
  
  ordensql = ordensql &      ")"

  orden_consulta = ordensql

  set objRecordset = OBJConnection.Execute(ordensql) 

 end if 
end if



'

' ------------------------------------------------------------------------------------------------------------------------------------------
'
'
'
'------------------------------------------------------------------------
'
'
' Al usuario(a) indicándole que acaba de recibir respuesta del la pregunta
'
'if pregunta<>"" then
' orden = "SELECT usuario FROM ECOINFORMAS_CONSULTAS WHERE idconsulta='"&idconsulta&"'"
' set objRecordset = OBJConnection.Execute(orden)
' if not objRecordset.eof then
'  idgente = objRecordset("usuario")
'  modelo  = "3"
'  enviado = enviar_mail()
' end if 
' 
'else
' Al asesor/es indicándole que se le acaba de asignar un asesoramiento
'
' orden = "SELECT EG.email AS email FROM ECOINFORMAS_GENTE EG "
' orden = orden & " LEFT JOIN ECOINFORMAS_TEM_ASE ET ON ET.asesor=EG.idgente "
' orden = orden & " WHERE EG.asesor=2 and ET.valor='"&tema_consulta&"'"
' set objRecordset = OBJConnection.Execute(orden)
 
' destinatarios = ""
' conec = ""
' do while not objRecordset.eof
'  destinatarios = destinatarios & conec & trim(objRecordset("email"))
'  conec = ","
'  objRecordset.movenext
' loop
' modelo = "2"
' enviado = enviar_mail()
' 
'end if 
'------------------------------------------------------------------------


'
function dame_provincia(codpro)
 orden = "select SV.valor as xvalor from ECOINFORMAS_VALORES EV LEFT JOIN SAAT_VALORES SV ON SV.grupo='013' and SV.subgrupo=EV.subgrupo WHERE EV.valor='"&codpro&"'"
 set objRecordset = OBJConnection.Execute(orden)
 if not objRecordset.eof then
  dame_provincia = objRecordset("xvalor")
 else
  dame_provincia = "" 
 end if
end function

function dame_territorio(codpro)
 orden = "select V2.valor as valor from SAAT_VALORES V1 LEFT JOIN SAAT_VALORES V2 ON V2.grupo='032' and V2.subgrupo=V1.desc2 WHERE V1.valor='"&codpro&"'"
 set objRecordset = OBJConnection.Execute(orden)
 if not objRecordset.eof then
  dame_territorio = objRecordset("valor")
 else
  dame_territorio = "" 
 end if
end function


%>

<!--#include file="asesora_enviar_email.asp"-->

<script>
 //alert ("Orden:<%=orden%>");
 //alert ("Destinatarios:<%=destinatarios%>");

 param = 'act_pag=<%=act_pag%>&estado_consulta=<%=estado_consulta_pri%>&tema_consulta_fil=<%=tema_consulta_fil%>';
 opener.location.href='asesora_paso1.asp?'+param;
 window.close();
</script>

<!--
Orden_empresa: <%=orden_empresa%><br>
Orden_trabajador: <%=orden_trabajador%><br>
Orden_consulta: <%=orden_consulta%><br>


eco_tra_nombre: <%=eco_tra_nombre%><br>
idconsulta: <%=idconsulta%><br>
eco_con_idusuario: <%=eco_con_idusuario%><br>

-->

</body>
</html>

<!--#include file="unquote.asp"-->