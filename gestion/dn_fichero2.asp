<%
'++++++CODIGO PARALELO A dn_sustancia; IMAGEN es el aqui el pdf, doc, ...), estructura_molecular se llama aqui archivo ++++++
'++++++los uploads se guardan tb con su id.extension, pero antecedidos de "_" ++++++
'++++++la gestion de los uploads cambia un poco, por ejemplo, se manda el nombre de fichero que ya existia, por si se quiere eliminar (para las imagenes, bastaba con intentar borrar id.gif, id.jpg, ahora pueden subir cualquier extension


Response.expires = 0
tiempo=2000
%>

<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<!-- #include file="freeaspupload.asp" -->

<%'Permite auditar las acciones%>
<!--#include file="spl_auditoria.inc.asp"-->

<%	
		dim hayimagen, advertencia, tipo, archivotemp

		id=request.querystring("id")

		'Set Upload = New FreeASPUpload
		'salvaimagen=SaveFiles() 'IMAGEN si devuelve vacio, no habia imagen; si devuelve KO, ha ocurrido un error (se genera advertencia); else, devuelve nombre del archivotemp						
		
		' Para que funcione en el servidor de ISTAS
		Set Upload = Server.CreateObject("Persits.Upload.1")
		salvaimagen = SaveFiles_xip() 
		
		'RECOGEMOS VALORES:
		
			'si no hay error, hacemos insert/update según corresponda
	
		if menserror="" then
		
			titulo=h(Upload.Form("titulo"))
			num_alternativa=h(Upload.Form("num_alternativa"))
			if num_alternativa="" then num_alternativa="NULL"
			tema=h(Upload.Form("tema"))
			resumen=h(Upload.Form("resumen"))
			direccion_internet=h(Upload.Form("direccion_internet"))
			idioma=h(Upload.Form("idioma"))
			autor=h(Upload.Form("autor"))
			lugar=h(Upload.Form("lugar"))
			publicacion=h(Upload.Form("publicacion"))
			coleccion=h(Upload.Form("coleccion"))
			descripcion_fisica=h(Upload.Form("descripcion_fisica"))
			numero_normalizado=h(Upload.Form("numero_normalizado"))
			notas=h(Upload.Form("notas"))
			soporte=h(Upload.Form("soporte"))
			fecha_actualizacion=h(Upload.Form("fecha_actualizacion"))
			fecha_consulta=h(Upload.Form("fecha_consulta"))
			criterios_aceptacion=h(Upload.Form("criterios_aceptacion"))
			alternativas_minimizacion_residuos=h(Upload.Form("alternativas_minimizacion_residuos"))
				
			'si nos pasan id, update, si no, insert			
			if id<>"" then			
				' ** AUDITORIA **
				spl_accion = "modificar"
				
				select case Upload.Form("imagen")
				
					case "nueva","cambiar": 'solo realizamos esta accion si nos han enviado un archivo
									if hayimagen then
										miextension=dimeextension(archivotemp)
										response.write miextension
										'vemos si la extension es valida (si no ha devuelto un numero)
										if isnumeric(miextension) then
											advertencia="El nombre de archivo no es válido. Compruebe que no hay más puntos que el que separa el nombre de la extensión (ej.: recomendacionesagricultura.pdf)"
										else
											NewFileName=dimenuevonombre(id,miextension)
											if NewFileName="KO" then
												advertencia="El archivo no es válido."
											else
												'ANTEPONEMOS "_" al id, para distinguir ficheros de estructuras moleculares (van en la misma carpeta)
												RenombrarArchivo "\estructuras\temp\" &archivotemp, "\estructuras\_" &NewFileName											
												camposupd= "archivo='_" &NewFileName& "', "
												'borramos antigua
												ficheroantiguo=Upload.Form("ficheroantiguo")
												borrarfichero "\estructuras\" &ficheroantiguo
											end if														
										end if																	
									end if
									'(en el caso de cambiar, deberiamos borrar la antigua, pero en realidad no es necesario, pq lo que hacemos es machacarla)
									
					case "eliminar": camposupd= "archivo='', "
									'eliminamos de disco
									ficheroantiguo=Upload.Form("ficheroantiguo")
									ficheroantiguo=Upload.Form("ficheroantiguo")
									borrarfichero "\estructuras\" &ficheroantiguo

				end select
			
			
			
				camposupd=camposupd& "titulo='" & titulo &  "', num_alternativa="  & num_alternativa &  ", tema='"  & tema &  "', resumen='"  & resumen &  "', direccion_internet='"  & direccion_internet &  "', idioma='"  & idioma &  "', autor='"  & autor &  "', lugar='"  & lugar &  "', publicacion='"  & publicacion &  "', coleccion='"  & coleccion &  "', descripcion_fisica='"  & descripcion_fisica &  "', numero_normalizado='"  & numero_normalizado &  "', notas='"  & notas &  "', soporte='"  & soporte &  "', criterios_aceptacion='"&criterios_aceptacion&"', alternativas_minimizacion_residuos='"&alternativas_minimizacion_residuos&"'"
				if fecha_consulta<>"" then 
					camposupd=camposupd& ", fecha_consulta='"  & fecha_consulta &  "'"
				else
					camposupd=camposupd& ", fecha_consulta=NULL"
				end if
				if fecha_actualizacion<>"" then
					camposupd=camposupd& ", fecha_actualizacion='"  & fecha_actualizacion &  "'"
				else
					camposupd=camposupd& ", fecha_actualizacion=NULL"
				end if
				
				sqlm="UPDATE dn_alter_ficheros SET " &camposupd& " WHERE id=" &id
'response.write sqlm
				'objconn1.Execute sqlm, lngRecs, adCmdText,, adExecuteNoRecords
' SE COMENTA PARA NO PERDER LOS DATOS
				objconn1.Execute sqlm, lngRecs, adExecuteNoRecords + adCmdText
				if lngRecs=0 then
					flashMsgCreate "No se ha encontrado en la base la información del fichero a modificar.", "Error"
					cerrarconexion
					response.redirect "dn_fichero.asp" 
				end if
				
			else 'nuevo:insert
			' ** AUDITORIA **
			spl_accion = "crear"
			
			lugar=h(Upload.Form("lugar"))
			publicacion=h(Upload.Form("publicacion"))
			coleccion=h(Upload.Form("coleccion"))
			descripcion_fisica=h(Upload.Form("descripcion_fisica"))
			numero_normalizado=h(Upload.Form("numero_normalizado"))
			notas=h(Upload.Form("notas"))
			soporte=h(Upload.Form("soporte"))
			fecha_actualizacion=h(Upload.Form("fecha_actualizacion"))
			fecha_consulta=h(Upload.Form("fecha_consulta"))
			criterios_aceptacion = h(Upload.Form("criterios_aceptacion"))
			
				campos="titulo, num_alternativa, tema, resumen, direccion_internet, idioma, autor, lugar, publicacion, coleccion, descripcion_fisica, numero_normalizado, notas, soporte, criterios_aceptacion, alternativas_minimizacion_residuos"
				if fecha_consulta<>"" then campos=campos& ", fecha_consulta" 
				if fecha_actualizacion<>"" then campos=campos& ", fecha_actualizacion" 
				
				valores="'" &titulo&"', "& num_alternativa&", '"& tema&"', '"& resumen&"', '"& direccion_internet&"', '"& idioma&"', '"& autor&"'"
				valores=valores& ", '"& lugar& "','" &publicacion&"','"&  coleccion&"','"&  descripcion_fisica&"','"&  numero_normalizado&"','"&  notas&"','"&  soporte&"','"&criterios_aceptacion&"'"
				valores=valores& ", '"& alternativas_minimizacion_residuos & "'"
				
				if fecha_consulta<>"" then valores=valores& ",'"&  fecha_consulta&"'"
				if fecha_actualizacion<>"" then valores=valores& ",'"&  fecha_actualizacion&"'"
				
					sqlm="INSERT INTO dn_alter_ficheros (" &campos& ") VALUES (" &valores& ")"
					response.write sqlm
					objconn1.execute(sqlm)
					'cogemos id
					set rstid=objconn1.execute("select top 1 id from dn_alter_ficheros order by id desc")
					id=rstid("id")
					rstid.close
					set rstid=nothing
					
					'renombramos imagen (si hay) y actualizamos base
					if hayimagen then
						miextension=dimeextension(archivotemp)
						'vemos si la extension es valida (si no ha devuelto un numero)
						if isnumeric(miextension) then
							advertencia="El nombre de archivo no es válido. Compruebe que no hay más puntos que el que separa el nombre de la extensión (ej.: recomendacionesagricultura.pdf)"
						else					
							NewFileName=dimenuevonombre(id,miextension)	
							if NewFileName="KO" then
								advertencia="La extensión del archivo no es válida. Sólo se admiten archivos gif y jpg."
							else
								RenombrarArchivo "\estructuras\temp\" &archivotemp, "\estructuras\_" &NewFileName										
								sqlm="UPDATE dn_alter_ficheros SET archivo='_" &NewFileName& "' WHERE id=" &id
								objconn1.execute(sqlm)
							end if
						end if
					end if
					
			end if 'if id		
			
		end if 'if menserror="" then
		
		'FINAL: si nos subieron imagen a temp, la borramos
		if hayimagen then borrarfichero "\estructuras\temp\" &archivotemp
		
' ** AUDITORIA **
spl_entidad = "fichero"
spl_descripcion = sqlm	
call auditaYCierraConexion(spl_accion,spl_entidad,spl_descripcion) ' accion, entidad, descripcion			
		
		'si ha habido error, advertimos; si no, cerramos ventana y actualizamos padre
		if menserror<>"" then
			mostrarmens
		else
			if advertencia<>"" then
				flashMsgCreate advertencia, "advertencia"
			else
				flashMsgCreate "La base de ficheros se ha actualizado con los nuevos datos.", "OK"
			end if
' SE COMENTA EL REDIRECCIONAMIENTO
			response.redirect "dn_fichero.asp?id=" &id
			
			'flashMsgShow			
		end if	

		Set Upload = nothing

%>

<%
function mostrarmens
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Istas</title>
<link rel="stylesheet" type="text/css" href="dn_estilos.css">
<script type="text/javascript" src="niftycube.js"></script>
<script type="text/javascript">
window.onload=function(){
Nifty("div#box2","big"); 
}
</script>
</head>

<body>
<%
flashMsgCreate menserror, "Error"
flashMsgShow()
%>
<!-- <div id="box2" class="centcontenido">
<%=menserror%>
<br><br> -->
<div align="center"><input type="button" value="Volver" onclick="history.back();" /></div>
<!-- </div> -->
</body>
</html>

<%
end function
%>
<%
function SaveFiles()
    
	on error resume next
	
    'directorio donde se salvan las imagenes
	uploadsDirVar = server.mappath(".") & "\estructuras\temp\"
    Upload.Save(uploadsDirVar)
	' If something fails inside the script, but the exception is handled
	If Err.Number <> 0 then
		resultado = "KO"
	Else
		ks = Upload.UploadedFiles.keys
		if (UBound(ks) <> -1) then		
			for each fileKey in Upload.UploadedFiles.keys
				resultado =  Upload.UploadedFiles(fileKey).FileName 
				tipo=Upload.UploadedFiles(fileKey).ContentType
			next
		else
			resultado = ""
		end if
	End if
	
	select case resultado
					case "": 	hayimagen=false
					case "KO": 	hayimagen=false
							  	advertencia="El archivo no se ha insertado, porque ha ocurrido un error."
					case else: 	hayimagen=true
								archivotemp=resultado					
	end select
	SaveFiles=resultado
	
end function


function SaveFiles_xip()

 	on error resume next

	ruta_upload_fis = server.mappath(".") & "\estructuras\temp\"

	Upload.ProgressID = Request.QueryString("PID")
	Upload.OverwriteFiles = false
	Upload.SetMaxSize 300000000, True
	Count = Upload.Save(ruta_upload_fis)

	if Err <> 0 Then
 	 resultado = "KO"
	else
		resultado = ""
		if count<>0 then
  		Set File = Upload.Files(1)
  		resultado = File.fileName
 		end if 
 	end if
 	
	select case resultado
		case "": 	hayimagen=false
		case "KO": 	hayimagen=false
					advertencia="El archivo no se ha insertado, porque ha ocurrido un error."
		case else: 	hayimagen=true
					archivotemp=resultado					
	end select
	SaveFiles_xip = resultado
	
end function
%>
