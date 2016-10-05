<%
		sqlpag="select id, nombre from dn_risc_sustancias as sus WHERE id IN(" &cadenaids& ") ORDER BY " &ordenacion&  " " &sentido
		'response.write sqlpag
		set rstpag=objConnection2.execute(sqlpag)
		if not rstpag.eof then
			arrayDatos = rstpag.GetRows

			for contadorFilas=0 to currentPageFinalRecordNumber-currentPageInitialRecordNumber

						  tablares=tablares & "<tr>"
						  select case  filtro
						  	case "1": enlazacon="dn_alternativas_ficha_sustancia.asp"
						  	case else:enlazacon="dn_risctox_ficha_sustancia.asp"
						  end select
						  'Sergio -> por si hay uno solo, lo cojo
						  unico_enlace = enlazacon& "?id_sustancia=" &arrayDatos(0,contadorFilas)
						  tablares=tablares & "<td class='celda_risctox'><a href='" &enlazacon& "?id_sustancia=" &arrayDatos(0,contadorFilas)& "'>" &corta(arrayDatos(1,contadorFilas),100, "puntossuspensivos")& "</a><br />" & dameSinonimos(arrayDatos(0,contadorFilas)) & dameNombreingles(arrayDatos(0,contadorFilas))& dameNombrecomercial(arrayDatos(0,contadorFilas)) & "</td>"
						  tablares=tablares & "</tr>"
					'end if
			next
		end if
		rstpag.close
		set rstpag=nothing

		tablares="<table class='tabla3' width='90%' align='center' border='0' cellpadding='4' cellspacing='0'>" &tablares& "</table>"
%>