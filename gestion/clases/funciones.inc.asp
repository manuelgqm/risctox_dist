<%
	function devuelve_nivel(valor)
		select case	valor
			case isNull(valor)
				devuelve_nivel = ""
			case 0
				devuelve_nivel = ""
			case 1
				devuelve_nivel = "Bajo"
			case 2
				devuelve_nivel = "Medio"
			case 3
				devuelve_nivel = "Alto"
		end select
	end function

	function generaCampoFecha(nombre_campo,valor,formato)

		if (len(formato)=0) then formato = "%d/%m/%Y %H:%M"
		result = strComp(formato,"%d/%m/%Y %H:%M")
		if result = 0 then
			if ((valor <> "") and (isDate(valor))) then valor = FormatDateTime(valor,VBShortDate)&" "&FormatDateTime(valor,VBShortTime)
			str = 		"<input name="""&nombre_campo&""" id="""&nombre_campo&""" size=""17"" value="""&valor&""">" & vbCrlf
		else
			str = 		"<input name="""&nombre_campo&""" id="""&nombre_campo&""" size=""9"" value="""&valor&""">" & vbCrlf
		end if
		str = str & "<input style=""border: 0px;"" class=""boton"" type=""image"" src=""imagenes/calendario.gif"" id=""f_trigger_"&nombre_campo&""" />" & vbCrlf
		str = str & "<script type=""text/javascript"">" & vbCrlf
		str = str & "	Calendar.setup({" & vbCrlf
		str = str & "		inputField     :    """&nombre_campo&""",     		// ID del campo de entrada de datos" & vbCrlf
		str = str & " 		ifFormat       :    """&formato&""",				// Formato del campo de entrada de datos" & vbCrlf
		if result = 0 then
			str = str & " 		showsTime      :    true,	          				// No se mostrara el selector de tiempo" & vbCrlf
			str = str & " 		timeFormat     :    ""24"",							// Formato para la visualizacion de la hora" & vbCrlf
		end if
		str = str & " 		button         :    ""f_trigger_"&nombre_campo&""", // Disparador para el calendario (ID del boton)" & vbCrlf
		str = str & " 		singleClick    :    false,          				// Modo Doble-click" & vbCrlf
		str = str & " 		step           :    1               				// Muestra todos los a�os en cajas desplegables" & vbCrlf
		str = str & " 		});" & vbCrlf
		str = str & "		</script>" & vbCrlf
		generaCampoFecha = str
	end function

	' Comprueba la validez del email
	' La validaci�n se realiza en 2 partes:
	' validaci�n del formato
	' validaci�n del servidor
	function validaEmail(email)
		msjErrorFormato = validaFormatoMail(email)
'		if (len(msjErrorFormato) = 0 ) then
'			partes = Split(email, "@")
'			msjErrorServidor =  validaServidorMail(partes(1))
'		end if
'		validaEmail = msjErrorFormato & msjErrorServidor
		validaEmail = msjErrorFormato
	end function

	' Comprueba que el dominio indicado posee servidor de correo mediante el registro MX (mail exchange)
	function validaServidorMail(dominio)
		Dim objXMLHTTP,strResult
		Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
		objXMLHTTP.Open "Get", "http://examples.softwaremodules.com/IntraDns.asp?domainname=" & dominio & "&Submit=Submit&t_mx=1", False
		objXMLHTTP.Send
		strResult = objXMLHTTP.ResponseText
		strResult = Mid(strResult,InStr(strResult,"(MX) for <strong>"),100)
		strResult = Mid(strResult,Instr(strResult,"</strong>. Items Returned: <strong>")+35,1)
		if IsNumeric(strResult) then
		 	if CInt(strResult) > 0 then
				validaServidorMail = ""
			else
				validaServidorMail = "<li>La direcci�n de correo no es correcta</li>"
			end if
		else
			validaServidorMail = "<li>La direcci�n de correo no es correcta</li>"
		end if
	end function

	' comprueba que el formato del email es correcto
	function validaFormatoMail(email)
		validaFormatoMail = ""
		dim partes, nombre, i, c
		' obtiene las partes que forman el email: nombre y dominio
		partes = Split(email, "@")
		' comprueba el n�mero de partes que posee el email indicado
		' si es distinto de 1 (no existe @, o existe m�s de una), el email no es correcto
		if UBound(partes) <> 1 then
			validaFormatoMail = "<li>El email no es correcto</li>"
			exit function
		end if

		' si se encuentran 2 puntos seguidos, el email no es correcto
		if InStr(email, "..") > 0 then
			validaFormatoMail = "<li>El email no es correcto</li>"
		end if

		' si en el nombre del dominio no existe un ., el email no es correcto
		if InStr(partes(1), ".") <= 0 then
			validaFormatoMail = "<li>El email no es correcto</li>"
			exit function
		end if

		' comprueba que la extensi�n del dominio no sea de 2 o 3 caracteres
		i = Len(partes(1)) - InStrRev(partes(1), ".")
		if i <> 2 and i <> 3 then
			validaFormatoMail = "<li>El email no es correcto</li>"
			exit function
		end if

		' se comprueba la validez del formato de cada una de las partes que conforman el email
		for each nombre in partes
			' si la longitud es 0 significa que el email no es correcto (@dominio.com o nombre@)
			if Len(nombre) <=  0 then
				validaFormatoMail = "<li>El email no es correcto</li>"
				exit function
			end if
			' comprueba que los caracteres sean correctos
			for i = 1 to Len(nombre)
				c = Lcase(Mid(nombre, i, 1))
				letras ="abcdefghijklmnopqrstuvwxyz_-."
				' si el caracter tratado no est� entre los v�lidos, el email no es correcto
				if InStr(letras, c) <= 0 and not IsNumeric(c) then
					validaFormatoMail = "<li>El email no es correcto</li>"
					exit function
				end if
			next
			' si posee un punto inicial o final, el email no es correcto
			if Left(nombre, 1) = "." or Right(nombre, 1) = "." then
				validaFormatoMail = "<li>El email no es correcto</li>"
				exit function
			end if
		next
	end function

	function getUltimoId()
		set rs=Server.CreateObject("ADODB.recordset")
		'obtiene el id insertado
		sql = "SELECT @@IDENTITY AS ultimo_id"
	  	on error resume next
		conn.Execute sql
		rs.Open sql, conn
		getUltimoId = rs.fields.item("ultimo_id")
		rs.Close
	end function

	' Elimina el contenido de un directorio antes de borrarlo
	sub eliminaDirectorio(nom_dir)
		' crea un objeto fichero
		set filesys=CreateObject("Scripting.FileSystemObject")
		' comprueba si existe el directorio que va a eliminar
		If filesys.FolderExists(nom_dir) Then
			set directorio = filesys.GetFolder(nom_dir)
			for each fichero in directorio.Files
				' elimina el fichero
				filesys.DeleteFile nom_dir&"/"&fichero.Name
			next
			' Finalmente elimina el directorio
			directorio.Delete
		End If
	end sub


	function addslashes(cadena)
		if (Len(cadena)>0) then
			cadena = Replace(cadena,"'","''")
		end if
		addslashes = cadena
	end function

	Function FechaGenerica(fecha)
    If IsDate(fecha) = True Then
       DIM dia, mes, ano, hora
       hora= FormatDateTime(fecha,3) 'hh:mm:ss
       dia = Day(fecha)
       mes = Month(fecha)
       ano = Year(fecha)
       FechaGenerica = ano & "" & Right(Cstr(mes + 100),2) & "" & Right(Cstr(dia + 100),2) & " " & hora
    Else
       FechaGenerica = Null
    End If
End Function

	function nl2br(cadena)
		if cadena <> Null then
			cadena = Replace(cadena,vbCrLf,"<br />")
		end if
		nl2br = cadena
	end function

	function muestraProgramaciones(nifSol,programacion,formulario)
		dim rs,lista, sql

		lista = ""

		set rs=Server.CreateObject("ADODB.recordset")
		sql = "SELECT * FROM "&PRE_APLI&"programacion P"
		sql = sql & " LEFT JOIN "&PRE_APLI&"solicitante S ON S.pro_id = P.pro_id"
		sql = sql & " WHERE sol_nif='"&nifSol&"'"
		rs.Open sql, conn, 3, 1

		select case formulario
			case "itinerario":
				lista = "<select name='select_programacion' onChange='javascript:recargaProgramaciones(this.value,"""&formulario&""")'>"&vbCrlf
			case "solicitante","balance":
				lista = "<select name='select_programacion' onChange='javascript:document.form_edita.accion.value=""guardar"";document.form_edita.submit();recargaProgramaciones(this.value,"""&formulario&""")'>"&vbCrlf
		end select

'		lista = lista & "<option value='0'> Seleccionar </option>"&vbCrlf
		do until rs.EOF
			lista = lista & "<option value='"&rs("sol_id")&"'"
			if (CInt(programacion)=CInt(trim(rs.fields.item("pro_id")))) then
				lista = lista & " selected"
			end if
			' Para el caso de que se trate de la tabla usuario y el perfil Supervisor, se muestran
			' todos los orientadores y el centro al que pertencen
			lista = lista & ">"&trim(rs.fields.item("pro_descripcion"))&"</option>"& vbCrlf
			rs.MoveNext
		loop
		lista = lista &"</select>"&vbCrlf
		muestraProgramaciones = lista
  	end function



  	function getListaAuxiliar(nombreTabla, campo, valor)
		dim rs,lista, sql
		lista=""
		set rs=Server.CreateObject("ADODB.recordset")
		sql = "SELECT * FROM "&PRE_APLI&nombreTabla&" ORDER BY 2 ASC"


		if (nombreTabla="centro") then
			sql = "SELECT cen_id,cen_nombre FROM "&PRE_APLI&nombreTabla&" ORDER BY cen_nombre ASC"
		end if

		if (nombreTabla="usuario") then
			select case (objUsuConectado.getProperty("usu_perfil"))
				case "O", "AV"
					sql = "PROTECCI�N"
					if (isNumeric(objUsuConectado.getProperty("cen_id")) and isNumeric(SESSION.Contents("usu_programacion"))) then
						sql = "SELECT usu_id,usu_nombre + ' ' + usu_apellidos as nombre_completo"
						sql = sql & " FROM "&PRE_APLI&nombreTabla
						sql = sql & " WHERE usu_perfil = 'O'"
						sql = sql & " AND usu_borrado=0 "
						sql = sql & " AND usu_id in "
						sql = sql & "(select usu_id from " & PRE_APLI &"rel_usuario_programacion where cen_id="& objUsuConectado.getProperty("cen_id")
						sql = sql & " and pro_id=" & SESSION.Contents("usu_programacion") & ")"
						sql = sql & " ORDER BY nombre_completo ASC"
					end if
				case else
					sql = "PROTECCI�N"
					if (IsNumeric(SESSION.Contents("usu_programacion"))) then
						sql = "SELECT U.usu_id, cen_nombre, usu_nombre + ' ' + usu_apellidos as nombre_completo"
						sql = sql & " FROM "&PRE_APLI&nombreTabla&" U ,"&PRE_APLI&"centro C, " & PRE_APLI & "rel_usuario_programacion P"
						sql = sql & " WHERE C.cen_id = P.cen_id"
						sql = sql & " AND U.usu_id = P.usu_id"
						sql = sql & " AND usu_borrado=0 "
						sql = sql & " AND usu_perfil = 'O'"
						sql = sql & " AND P.pro_id = " & SESSION.Contents("usu_programacion")
						sql = sql & " ORDER BY cen_nombre ASC, nombre_completo ASC"
					end if
			end select
		end if

'la que habia en accion_funciones
'		if (nombreTabla="usuario") then
'			select case (objUsuConectado.getProperty("usu_perfil"))
'				case "O", "AV"
'					sql = " SELECT U.usu_id,usu_nombre + ' ' +usu_apellidos as nombre_completo,cen_nombre"
'					sql = sql & " FROM "&PRE_APLI&"usuario U, "&PRE_APLI&"rel_usuario_programacion R, "&PRE_APLI&"centro C"
'					sql = sql & " WHERE R.usu_id = U.usu_id"
'					sql = sql & " AND C.cen_id = R.cen_id "
'					sql = sql & " AND usu_perfil='O' "
'					sql = sql & " AND R.pro_id="&Session.Contents("usu_programacion")
'					sql = sql & " AND R.cen_id="&objUsuConectado.getProperty("cen_id")
'					sql = sql & " ORDER BY nombre_completo ASC"
'				case else
'					sql = " SELECT U.usu_id,usu_nombre + ' ' + usu_apellidos as nombre_completo,cen_nombre"
'					sql = sql & " FROM "&PRE_APLI&"usuario U, "&PRE_APLI&"rel_usuario_programacion R, "&PRE_APLI&"centro C"
'					sql = sql & " WHERE R.usu_id = U.usu_id"
'					sql = sql & " AND C.cen_id = R.cen_id "
'					sql = sql & " AND R.pro_id="&Session.Contents("usu_programacion")
'					sql = sql & " AND usu_perfil='O'"
'					sql = sql & " ORDER BY cen_nombre ASC, nombre_completo ASC"
'			end select
'		end if

		if (nombreTabla = "aux_frecuencia_consulta_mail") then
			sql = "SELECT * FROM "&PRE_APLI&nombreTabla
		end if

		if (nombreTabla = "aux_tiempo_cotizado") then
			sql = "SELECT * FROM "&PRE_APLI&nombreTabla
		end if

		if (nombreTabla = "aux_tiempo_busqueda") then
			sql = "SELECT * FROM "&PRE_APLI&nombreTabla
		end if

		if (nombreTabla = "aux_ambito_geografico_disp") then
			sql = "SELECT * FROM "&PRE_APLI&nombreTabla
		end if

		if (nombreTabla = "aux_duracion_curso") then
			sql = "SELECT * FROM "&PRE_APLI&nombreTabla
		end if

		if (nombreTabla = "aux_solicitante_estado") then
			sql = "SELECT * FROM "&PRE_APLI&nombreTabla
		end if

		if (nombreTabla="aux_estado_hoja_firmas") then
			sql = "SELECT * FROM "&PRE_APLI&nombreTabla&" ORDER BY 1 ASC"
		end if

		if (nombreTabla="aux_tipo_accion_acciones_grupales") then
			sql = "SELECT * FROM "&PRE_APLI&"aux_tipo_accion"
			sql = sql & " WHERE xta_tipo='G'"
			sql = sql & " ORDER BY 1 ASC"
		end if

		if (nombreTabla="aux_tipo_accion_acciones_individuales") then
			sql = "SELECT * FROM "&PRE_APLI&"aux_tipo_accion"
			sql = sql & " WHERE xta_tipo='I'"
			sql = sql & " ORDER BY 1 ASC"
		end if

		rs.Open sql, conn, 3, 1

		lista = "<select name='"&campo&"'>"'&vbCrlf
		lista = lista & "<option value='0'> Seleccionar </option>"
		do until rs.EOF
			lista = lista & "<option value='"&rs.fields.item(0)&"'"
			if (trim(valor)=trim(rs.fields.item(0))) then
				lista = lista & " selected"
			end if
			' Para el caso de que se trate de la tabla usuario y el perfil Supervisor, se muestran
			' todos los orientadores y el centro al que pertencen
			if ((nombreTabla = "usuario") and (objUsuConectado.getProperty("usu_perfil") = "S")) then
				lista = lista & ">"&trim(rs.fields.item(1))&" - "&trim(rs.fields.item(2))&"</option>" '& vbCrlf
				rs.MoveNext
			else
				lista = lista & ">"&trim(rs.fields.item(1))&"</option>" '& vbCrlf
				rs.MoveNext
			end if
		loop
		lista = lista &"</select>"'&vbCrlf
		getListaAuxiliar = lista
  	end function



  	function getListaCerrada(diccionario, campo, valor)
		dim lista,i
		lista=""
		claves = diccionario.keys
		lista = "<select name='"&campo&"'>"'&vbCrlf
		lista = lista & "<option value='0'> Seleccionar</option>"
		for i=0 to ubound(claves)
			lista = lista & "<option value='"&claves(i)&"'"
			if (trim(valor)=claves(i)) then
				lista = lista & " selected"
			end if
			lista = lista & ">"&diccionario.item(claves(i))&"</option>" '& vbCrlf
		next

		lista = lista &"</select>"'&vbCrlf
		getListaCerrada = lista
  	end function



  	function getInput(campo, valor, tipo, clase)
		resultado = "<input type='"&tipo&"' name='" & campo & "' value='" & valor & "' class='"&clase&"'>"
		getInput = resultado
  	end function






%>