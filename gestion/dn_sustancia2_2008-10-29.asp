<%
Response.expires = 0
tiempo=2000
%>

<!--#include file="dn_fun_comunes.asp"-->
<!--#include file="dn_fun_texto.asp"-->
<!--#include file="dn_auten.inc"-->
<!--#include file="adovbs.inc"-->
<!--#include file="dn_conexion.asp"-->
<!-- #include file="freeaspupload.asp" -->

<%	
		dim hayimagen, advertencia, tipo, archivotemp

		id=request.querystring("id")

		Set Upload = New FreeASPUpload
		salvaimagen=SaveFiles() 'IMAGEN si devuelve vacio, no habia imagen; si devuelve KO, ha ocurrido un error (se genera advertencia); else, devuelve nombre del archivotemp						
		
		'RECOGEMOS VALORES:
	
		'lo primero, validamos sinonimos y nombres_comerciales, y los metemos en array para luego
		sinonimos=Upload.Form("sinonimos")
		if sinonimos<>"" then
			sinonimos=replace(sinonimos,"@ ","@")
			sinonimos=replace(sinonimos," @","@")
			arrsinonimos=split(sinonimos, "@")					
			FOR i=0 to UBound(arrsinonimos)		
				menserror=menserror&comprobarl(arrsinonimos(i),750,"sinónimos")
			NEXT			
		end if
		nombres_comerciales=Upload.Form("nombres_comerciales")
		if nombres_comerciales<>"" then
			nombres_comerciales=replace(nombres_comerciales,"@ ","@")
			nombres_comerciales=replace(nombres_comerciales," @","@")
			arrnombres_comerciales=split(nombres_comerciales, "@")					
			FOR i=0 to UBound(arrnombres_comerciales)		
				menserror=menserror&comprobarl(arrnombres_comerciales(i),100,"nombres comerciales")
			NEXT			
		end if
		
			'si no hay error, hacemos insert/update según corresponda
	
		if menserror="" then
		
			nombre=h(Upload.Form("nombre"))
			nombre_ing=h(Upload.Form("nombre_ing"))
			num_rd=h(Upload.Form("num_rd"))
			num_ce_einecs=h(Upload.Form("num_ce_einecs"))
			num_ce_elincs=h(Upload.Form("num_ce_elincs"))
			num_cas=h(Upload.Form("num_cas"))
			num_onu=h(Upload.Form("num_onu"))
			formula_molecular=h(Upload.Form("formula_molecular"))
			estructura_molecular=h(Upload.Form("estructura_molecular"))
			simbolos=h(Upload.Form("simbolos"))
			clasificacion_1=h(Upload.Form("clasificacion_1"))
			clasificacion_2=h(Upload.Form("clasificacion_2"))
			clasificacion_3=h(Upload.Form("clasificacion_3"))
			clasificacion_4=h(Upload.Form("clasificacion_4"))
			clasificacion_5=h(Upload.Form("clasificacion_5"))
			clasificacion_6=h(Upload.Form("clasificacion_6"))
			clasificacion_7=h(Upload.Form("clasificacion_7"))
			clasificacion_8=h(Upload.Form("clasificacion_8"))
			clasificacion_9=h(Upload.Form("clasificacion_9"))
			clasificacion_10=h(Upload.Form("clasificacion_10"))
			clasificacion_11=h(Upload.Form("clasificacion_11"))
			clasificacion_12=h(Upload.Form("clasificacion_12"))
			clasificacion_13=h(Upload.Form("clasificacion_13"))
			clasificacion_14=h(Upload.Form("clasificacion_14"))
			clasificacion_15=h(Upload.Form("clasificacion_15"))
			frases_s=h(Upload.Form("frases_s"))
			conc_1=h(Upload.Form("conc_1"))
			eti_conc_1=h(Upload.Form("eti_conc_1"))
			conc_2=h(Upload.Form("conc_2"))
			eti_conc_2=h(Upload.Form("eti_conc_2"))
			conc_3=h(Upload.Form("conc_3"))
			eti_conc_3=h(Upload.Form("eti_conc_3"))
			conc_4=h(Upload.Form("conc_4"))
			eti_conc_4=h(Upload.Form("eti_conc_4"))
			conc_5=h(Upload.Form("conc_5"))
			eti_conc_5=h(Upload.Form("eti_conc_5"))
			conc_6=h(Upload.Form("conc_6"))
			eti_conc_6=h(Upload.Form("eti_conc_6"))
			conc_7=h(Upload.Form("conc_7"))
			eti_conc_7=h(Upload.Form("eti_conc_7"))
			conc_8=h(Upload.Form("conc_8"))
			eti_conc_8=h(Upload.Form("eti_conc_8"))
			conc_9=h(Upload.Form("conc_9"))
			eti_conc_9=h(Upload.Form("eti_conc_9"))
			conc_10=h(Upload.Form("conc_10"))
			eti_conc_10=h(Upload.Form("eti_conc_10"))
			conc_11=h(Upload.Form("conc_11"))
			eti_conc_11=h(Upload.Form("eti_conc_11"))
			conc_12=h(Upload.Form("conc_12"))
			eti_conc_12=h(Upload.Form("eti_conc_12"))
			conc_13=h(Upload.Form("conc_13"))
			eti_conc_13=h(Upload.Form("eti_conc_13"))
			conc_14=h(Upload.Form("conc_14"))
			eti_conc_14=h(Upload.Form("eti_conc_14"))
			conc_15=h(Upload.Form("conc_15"))
			eti_conc_15=h(Upload.Form("eti_conc_15"))
			notas_rd_363=h(Upload.Form("notas_rd_363"))
			notas_rd_363 = mid(notas_rd_363,1,199)
			notas_xml=h(Upload.Form("notas_xml"))

			frases_r_danesa=h(Upload.Form("frases_r_danesa"))								
				
			'si nos pasan id de sustancia, update, si no, insert			
			if id<>"" then			
				
				select case Upload.Form("imagen")
				
					case "nueva","cambiar": 'solo realizamos esta accion si nos han enviado un archivo
									if hayimagen then
										
										select case tipo
											case "image/jpeg": extension="jpg"
											case "image/gif": extension="gif"
											case else: extension="?"
										end select
										NewFileName=dimenuevonombre(id,extension)	
										if NewFileName="KO" then
											advertencia="La extensión del archivo no es válida. Sólo se admiten archivos gif y jpg."
										else
											RenombrarArchivo "\estructuras\temp\" &archivotemp, "\estructuras\" &NewFileName											
											camposupd= "estructura_molecular='" &NewFileName& "', "
										end if										
									end if
									'(en el caso de cambiar, deberiamos borrar la antigua, pero en realidad no es necesario, pq lo que hacemos es machacarla)
									
					case "eliminar": camposupd= "estructura_molecular='', "
									'eliminamos de disco
									borrarfichero "\estructuras\" &id& ".gif" 
									borrarfichero "\estructuras\" &id& ".jpg" 

				end select
				
				camposupd=camposupd& "nombre='" & nombre &  "', nombre_ing='"  & nombre_ing &  "', num_rd='"  & num_rd &  "', num_ce_einecs='"  & num_ce_einecs &  "', num_ce_elincs='"  & num_ce_elincs &  "', num_cas='"  & num_cas &  "', num_onu='"  & num_onu &  "', formula_molecular='"  & formula_molecular &  "', simbolos='"  & simbolos &"', clasificacion_1='"  & clasificacion_1 &  "', clasificacion_2='"  & clasificacion_2 &  "', clasificacion_3='"  & clasificacion_3 &  "', clasificacion_4='"  & clasificacion_4 &  "', clasificacion_5='"  & clasificacion_5 &  "', clasificacion_6='"  & clasificacion_6 &  "', clasificacion_7='"  & clasificacion_7 &  "', clasificacion_8='"  & clasificacion_8 &  "', clasificacion_9='"  & clasificacion_9 &  "', clasificacion_10='"  & clasificacion_10 &  "', clasificacion_11='"  & clasificacion_11 &  "', clasificacion_12='"  & clasificacion_12 &  "', clasificacion_13='"  & clasificacion_13 &  "', clasificacion_14='"  & clasificacion_14 &  "', clasificacion_15='"  & clasificacion_15 &  "', frases_s='"  & frases_s &  "', conc_1='"  & conc_1 &  "', eti_conc_1='"  & eti_conc_1 &  "', conc_2='"  & conc_2 &  "', eti_conc_2='"  & eti_conc_2 &  "', conc_3='"  & conc_3 &  "', eti_conc_3='"  & eti_conc_3 &  "', conc_4='"  & conc_4 &  "', eti_conc_4='"  & eti_conc_4 &  "', conc_5='"  & conc_5 &  "', eti_conc_5='"  & eti_conc_5 &  "', conc_6='"  & conc_6 &  "', eti_conc_6='"  & eti_conc_6 &  "', conc_7='"  & conc_7 &  "', eti_conc_7='"  & eti_conc_7 &  "', conc_8='"  & conc_8 &  "', eti_conc_8='"  & eti_conc_8 &  "', conc_9='"  & conc_9 &  "', eti_conc_9='"  & eti_conc_9 &  "', conc_10='"  & conc_10 &  "', eti_conc_10='"  & eti_conc_10 &  "', conc_11='"  & conc_11 &  "', eti_conc_11='"  & eti_conc_11 &  "', conc_12='"  & conc_12 &  "', eti_conc_12='"  & eti_conc_12 &  "', conc_13='"  & conc_13 &  "', eti_conc_13='"  & eti_conc_13 &  "', conc_14='"  & conc_14 &  "', eti_conc_14='"  & eti_conc_14 &  "', conc_15='"  & conc_15 &  "', eti_conc_15='"  & eti_conc_15 &  "', notas_rd_363='"  & notas_rd_363 &  "', notas_xml='"  & notas_xml & "', frases_r_danesa='" & frases_r_danesa & "'"
				
				sqlm="UPDATE dn_risc_sustancias SET " &camposupd& " WHERE id=" &id
				'response.write sqlm
				'objconn1.Execute sqlm, lngRecs, adCmdText,, adExecuteNoRecords
				objconn1.Execute sqlm, lngRecs, adExecuteNoRecords + adCmdText
				if lngRecs=0 then
					flashMsgCreate "No se ha encontrado la sustancia a modificar.", "Error"
					cerrarconexion
					response.redirect "dn_sustancia.asp"
				end if
				
			else 'nuevo:insert
				campos="nombre, nombre_ing, num_rd, num_ce_einecs, num_ce_elincs, num_cas, num_onu, formula_molecular, simbolos, clasificacion_1, clasificacion_2, clasificacion_3, clasificacion_4, clasificacion_5, clasificacion_6, clasificacion_7, clasificacion_8, clasificacion_9, clasificacion_10, clasificacion_11, clasificacion_12, clasificacion_13, clasificacion_14, clasificacion_15, frases_s, conc_1, eti_conc_1, conc_2, eti_conc_2, conc_3, eti_conc_3, conc_4, eti_conc_4, conc_5, eti_conc_5, conc_6, eti_conc_6, conc_7, eti_conc_7, conc_8, eti_conc_8, conc_9, eti_conc_9, conc_10, eti_conc_10, conc_11, eti_conc_11, conc_12, eti_conc_12, conc_13, eti_conc_13, conc_14, eti_conc_14, conc_15, eti_conc_15, notas_rd_363, notas_xml, frases_r_danesa"
				valores="'" &nombre&"', '"& nombre_ing&"', '"& num_rd&"', '"& num_ce_einecs&"', '"& num_ce_elincs&"', '"& num_cas&"', '"& num_onu&"', '"& formula_molecular&"', '"& simbolos&"', '"& clasificacion_1&"', '"& clasificacion_2&"', '"& clasificacion_3&"', '"& clasificacion_4&"', '"& clasificacion_5&"', '"& clasificacion_6&"', '"& clasificacion_7&"', '"& clasificacion_8&"', '"& clasificacion_9&"', '"& clasificacion_10&"', '"& clasificacion_11&"', '"& clasificacion_12&"', '"& clasificacion_13&"', '"& clasificacion_14&"', '"& clasificacion_15&"', '"& frases_s&"', '"& conc_1&"', '"& eti_conc_1&"', '"& conc_2&"', '"& eti_conc_2&"', '"& conc_3&"', '"& eti_conc_3&"', '"& conc_4&"', '"& eti_conc_4&"', '"& conc_5&"', '"& eti_conc_5&"', '"& conc_6&"', '"& eti_conc_6&"', '"& conc_7&"', '"& eti_conc_7&"', '"& conc_8&"', '"& eti_conc_8&"', '"& conc_9&"', '"& eti_conc_9&"', '"& conc_10&"', '"& eti_conc_10&"', '"& conc_11&"', '"& eti_conc_11&"', '"& conc_12&"', '"& eti_conc_12&"', '"& conc_13&"', '"& eti_conc_13&"', '"& conc_14&"', '"& eti_conc_14&"', '"& conc_15&"', '"& eti_conc_15&"', '"& notas_rd_363&"', '"& notas_xml& "', '" & frases_r_danesa & "'"
					sqlm="INSERT INTO dn_risc_sustancias (" &campos& ") VALUES (" &valores& ")"
					'response.write sqlm
					objconn1.execute(sqlm)
					'cogemos id
					set rstid=objconn1.execute("select top 1 id from dn_risc_sustancias order by id desc")
					id=rstid("id")
					rstid.close
					set rstid=nothing
					
					'renombramos imagen (si hay) y actualizamos base
					if hayimagen then
						select case tipo
											case "image/jpeg": extension="jpg"
											case "image/gif": extension="gif"
											case else: extension="?"
						end select
						NewFileName=dimenuevonombre(id,extension)		
						if NewFileName="KO" then
							advertencia="La extensión del archivo no es válida. Sólo se admiten archivos gif y jpg. Se ha insertado la sustancia, pero no la imagen."
						else
							RenombrarArchivo "\estructuras\temp\" &archivotemp, "\estructuras\" &NewFileName										
							sqlm="UPDATE dn_risc_sustancias SET estructura_molecular='" &id&extension& "' WHERE id=" &id
							objconn1.execute(sqlm)
						end if
					end if
					
			end if 'if id		
			
		end if 'if menserror="" then
	
		'PARA INSERT Y UPDATE: borramos/insertamos de nuevo sinonimos y nombres comerciales

    objconn1.execute("delete from dn_risc_sinonimos where id_sustancia=" &id)
		if sinonimos<>"" then	
			FOR i=0 to UBound(arrsinonimos)		
				misinonimo=h(trim(arrsinonimos(i)))
				if misinonimo<>"" then objconn1.execute("insert into dn_risc_sinonimos (id_sustancia,nombre) values (" &id& ",'" &misinonimo&  "')")
			NEXT	
		end if
		
		'borramos/insertamos de nuevo nombres_comerciales y nombres comerciales
		objconn1.execute("delete from dn_risc_nombres_comerciales where id_sustancia=" &id)

		if nombres_comerciales<>"" then
			FOR i=0 to UBound(arrnombres_comerciales)	
				minc=trim(arrnombres_comerciales(i))
				if minc<>"" then objconn1.execute("insert into dn_risc_nombres_comerciales (id_sustancia,nombre) values (" &id& ",'" &minc&  "')")
			NEXT	
		end if
		
		'FINAL: si nos subieron imagen a temp, la borramos
		if hayimagen then borrarfichero "\estructuras\temp\" &archivotemp
		
		'si ha habido error, advertimos; si no, cerramos ventana y actualizamos padre
		if menserror<>"" then
			mostrarmens
		else
			if advertencia<>"" then
				flashMsgCreate advertencia, "advertencia"
			else
				flashMsgCreate "La base de sustancias se ha actualizado con los nuevos datos.", "OK"
			end if
			response.redirect "dn_sustancia.asp?id=" &id
			
			'flashMsgShow			
		end if	

		Set Upload = nothing

cerrarconexion
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
							  	advertencia="La imagen no se ha insertado, porque ha ocurrido un error."
					case else: 	hayimagen=true
								archivotemp=resultado					
	end select
	SaveFiles=resultado
	
end function
%>



