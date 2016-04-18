<%
ordenacion=request("ordenacion")
sentido=request("sentido") 
nregs=request("nregs") 

'valores de busqueda por defecto

if ordenacion="" then ordenacion="sus.nombre"
if sentido="" then sentido=""
if nregs="" then nregs=50
	
if busc="" then
	
else
	
	if isnumeric(nregs) then
		nregs=round(nregs,0)
	else
		nregs=50
	end if
	
	nombre=lcase(request.form("nombre"))
	tipobus=request.form("tipobus")
	numero=request.form("numero")
		
	select case busc
	
	case 1: 'han dado a buscar
			
		condicion=""	
		
			if nombre<>"" or numero<>"" then
			condicion=""
			if nombre<>"" then	'busca en nombre, sinonimos, nombre ingles y nombre comercial
				nombre2=h(nombre)	
				nombre2=quitartildes(nombre2)
				nombre2=montartildes(nombre2)
				if tipobus="exacto" then
					' La busqueda exacta tambien usa like, sin %, para no distinguir mayusculas

          ' CONDICION ANTIGUA
					'condicion=condicion& " (sus.nombre like'" &nombre2& "' or sin.nombre like'" &nombre2& "' or sus.nombre_ing like '" &nombre2& "' or com.nombre like '" &nombre2& "')  "
      
          ' CONDICION NUEVA
          ' Para que encuentre por nombre ingles en busqueda exacta hay que incluir los casos:
          ' "nom" (exacta)
          ' "nom@%" (al principio y seguido por otros)
          ' "%@nom" (al final y antecedido por otros)
          ' "%@nom@%" (en medio, seguido y precedido por otros)
          ' No debe haber espacio junto a las @ (avisar al cliente)

          condicion=condicion& " (sus.nombre like'" &nombre2& "' or sin.nombre like'" &nombre2& "' or sus.nombre_ing like '" &nombre2& "' or sus.nombre_ing like '" &nombre2& "@%' or sus.nombre_ing like '%@" &nombre2& "' or sus.nombre_ing like '%@" &nombre2& "@%' or com.nombre like '" &nombre2& "')  "  

				else
					condicion=condicion& " (sus.nombre like '%" &nombre2& "%' or sin.nombre like '%" &nombre2& "%' or sus.nombre_ing like '%" &nombre2& "%' or com.nombre like '%" &nombre2& "%')  "
				end if
			end if
			if numero<>"" then
				if nombre<>"" then condicion=condicion& " OR "
				condicion=condicion& " (num_ce_einecs = '" &numero& "' OR num_ce_elincs  = '" &numero& "' OR  num_rd = '" &numero& "' OR  num_cas = '" &numero& "')"
			end if
		end if
		
		sqls="select distinct sus.id, sus.nombre "
		'sqls=sqls & " , dn_risc_sustancias_por_usos.toxico, dn_alter_ficheros_por_sustancias.id_fichero, dn_risc_sustancias_por_grupos.id_grupo, dn_risc_grupos_por_usos.toxico, dn_alter_ficheros_por_grupos.id_fichero "
		sqls=sqls & " from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) "
		sqls=sqls & " FULL OUTER JOIN dn_risc_nombres_comerciales as com ON (sus.id=com.id_sustancia) " 
		
		'según filtro, unimos a distintas tablas
		if filtro<>"0" then
		
			select case filtro
				
				case "1":
					' Buscador de alternativas. Muestra las sustancias asociadas como toxicas a un uso para el que hay alternativas no toxicas,
					' y también sustancias asociadas a ficheros de alternativas
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_por_usos ON (sus.id=dn_risc_sustancias_por_usos.id_sustancia)"
				sqls=sqls & " LEFT OUTER JOIN dn_alter_ficheros_por_sustancias ON (sus.id=dn_alter_ficheros_por_sustancias.id_sustancia)"
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_por_grupos ON (sus.id=dn_risc_sustancias_por_grupos.id_sustancia)"
				sqls=sqls & " LEFT OUTER JOIN dn_risc_grupos_por_usos ON (dn_risc_sustancias_por_grupos.id_grupo=dn_risc_grupos_por_usos.id_grupo)"
				sqls=sqls & " LEFT OUTER JOIN dn_alter_ficheros_por_grupos ON (dn_risc_sustancias_por_grupos.id_grupo=dn_alter_ficheros_por_grupos.id_grupo)"
				
				case "cym": 'cancerigenos y mutagenos segun RD 363/1995
				'no unimos a mas tablas
				
				case "cym2": 'cancerigenos y mutagenos segun IARC
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia)"
				
				case "cym3": 'cancerigenos y mutagenos segun otras
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia)"
				
				case "cym3": 'cancerigenos y mutagenos segun otras
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia)"

				case "mama": 'cáncer de mama
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia)"

				case "cop": 'cop
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia)"
				
				case "tpr": 	'T&oacute;xicos para la reproducci&oacute;n
				
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_iarc iarc ON sus.id = iarc.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_cancer_otras otras ON sus.id = otras.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor neuro ON sus.id = neuro.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_ambiente ambiente ON sus.id = ambiente.id_sustancia " 
				
				case "dis": 'disruptor endocrino
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia)"
				
				case "neu": 'neurotoxico
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia)"

				case "sen": ' sensibilizante... no hace falta más tablas que la principal
				
				case "pyb": 'Sustancias tóxicas, persistentes y bioacumulativas
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				
				case "tac": 'Sustancias de toxicidad acuática según Directiva de aguas
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				
				case "tac2": 'Sustancias peligrosas agua Alemania
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				
				case "dat": 'Sustancias de daño a la capa de ozono
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				
				case "dat2": 'Sustancias cambio climático
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				
				case "dat3": 'Sustancias calidad aire
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				
				case "vl1": 'Límites de exposición profesional: Valores Límite Ambientales
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia)"
				
				case "vl2": '	Valores Límite Ambientales Cancerígenos
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_vl as vl ON (sus.id=vl.id_sustancia)"
				
				case "vl3": 'Límites de exposición profesional: Valores Límite Biológicos
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia)"
				
				case "enf": 'Enfermedades profesionales (borrador)
				sqls=sqls & " INNER JOIN dn_risc_sustancias_por_grupos ON (sus.id=dn_risc_sustancias_por_grupos.id_sustancia) INNER JOIN dn_risc_grupos_por_enfermedades ON (dn_risc_sustancias_por_grupos.id_grupo=dn_risc_grupos_por_enfermedades.id_grupo)"
				
				case "res": 'residuos
				sqls=sqls & " FULL OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				sqls=sqls & " FULL OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia)"
				sqls=sqls & " FULL OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia)"
				sqls=sqls & " FULL OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia)"
				sqls=sqls & " FULL OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia)"
				
				case "ver": 'vertidos
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_iarc iarc ON sus.id = iarc.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_cancer_otras otras ON sus.id = otras.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor neuro ON sus.id = neuro.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_ambiente ambiente ON sus.id = ambiente.id_sustancia "
				
				case "emi": 'emisiones
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				
				case "cov": 'Compuestos orgánicos volátiles (COV)
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				
				case "lpc": 'Sustancias (LPCIC)
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				
				case "acm": 'Sustancias que pueden provocar accidentes graves
				sqls=sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
				
				case "neg": ' Lista negra
				'no unimos a mas tablas
			end select
			
		end if
		
		if condicion<>"" then sqls=sqls & " WHERE (" &condicion& ")"
		
		'según filtro, agregamos distintas condiciones
		if filtro<>"0" then
			
			if condicion="" then
				sqls=sqls & " WHERE "
			else
				sqls=sqls & " AND "
			end if
			
			select case filtro
		
				case "1": 'la sustancia (o el grupo al que pertenece) debe tener un uso toxico, o debe existir una alternativa	 
				  sqls=sqls & " (dn_risc_sustancias_por_usos.toxico=1 OR dn_alter_ficheros_por_sustancias.id_fichero is not null OR dn_risc_grupos_por_usos.toxico=1 OR dn_alter_ficheros_por_grupos.id_fichero is not null) " 
				
 				case "cym":
          campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
          frases="R40, R45, R49, R40/20, R40/21, R40/22, R40/20/21, R40/20/22, R40/21/22, R40/20/21/22, R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/20/22, R68/21/22, R68/20/21/22"

				  sqls=sqls & "( " & monta_condicion(campos, frases) & " OR ("&monta_condicion_grupo("asoc_cancer_rd")&") )"
				
				case "cym2": 'la condicion simplemente es que exista la fila, pero ya que estamos, comprobamos que no este vacía
				  sqls=sqls & "( ((dn_risc_sustancias_iarc.grupo_iarc<>'')) OR ("&monta_condicion_grupo("asoc_cancer_iarc")&") )" 
				
				case "cym3": 'la condicion simplemente es que exista la fila, pero ya que estamos, comprobamos que no este vacía
				  sqls=sqls & "( ((dn_risc_sustancias_cancer_otras.categoria_cancer_otras<>'')) OR ("&monta_condicion_grupo("asoc_cancer_otras")&") )" 

				case "mama": 'que cancer_mama sea 1
				  sqls=sqls & "( ((dn_risc_sustancias_mama_cop.cancer_mama=1)) OR ("&monta_condicion_grupo("asoc_cancer_mama")&") )" 
				
				case "cop": 'que cop no sea vacío
				  sqls=sqls & "( ((dn_risc_sustancias_mama_cop.cop <> '')) OR ("&monta_condicion_grupo("asoc_cop")&") )"

				case "tpr":
          ' Buscando las frases R: TPR R60, R61, R62, R63, en las columnas CLASIFICACION_1, hasta CLASIFICACION_6 del RD 363/1995.

          campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
          frases="R60, R61, R62, R63"

				  sqls=sqls & "( " & monta_condicion(campos, frases) & " OR ("&monta_condicion_grupo("asoc_reproduccion")&") )"
				
				case "dis": 'nivel_disruptor no esta vacio
				  sqls=sqls & "( ((dn_risc_sustancias_neuro_disruptor.nivel_disruptor<>'')) OR ("&monta_condicion_grupo("asoc_disruptores")&") )" 
				
				case "neu": 'en campos adicionales nivel_neurotoxico no esta vacio
				  'sqls=sqls & " ((dn_risc_sustancias_neuro_disruptor.nivel_neurotoxico<>'')) " 
          sqls = sqls & "( " & sql_lista_neurotoxico & " OR ("&monta_condicion_grupo("asoc_neuro_oto")&") )"

				case "sen": 'determinadas frases R en clasif_1 a clasif_15 y frases_r_danesa
          campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15, sus.frases_r_danesa"
          frases = "R42, R43, R42/43, R42-43"
          sqls=sqls & monta_condicion(campos, frases)
				
				case "pyb": 
				  sqls=sqls & " ((dn_risc_sustancias_ambiente.anchor_tpb<>'')) " 
				
				case "tac": 
				  sqls=sqls & "( ((dn_risc_sustancias_ambiente.directiva_aguas=1)) OR ("&monta_condicion_grupo("asoc_directiva_aguas")&") )" 
				
				case "tac2": 
				  sqls=sqls & " ((dn_risc_sustancias_ambiente.clasif_MMA<>'')) " 
				
				case "dat": 
				  sqls=sqls & " ((dn_risc_sustancias_ambiente.dano_ozono=1)) " 
				
				case "dat2": 
				  sqls=sqls & " ((dn_risc_sustancias_ambiente.dano_cambio_clima=1)) " 
				
				case "dat3": 
				  sqls=sqls & "( ((dn_risc_sustancias_ambiente.dano_calidad_aire=1)) OR ("&monta_condicion_grupo("asoc_calidad_aire")&") )" 
				
				case "vl1": 
				  sqls=sqls & "( ((vla_ed_ppm_1<>'') or (vla_ed_mg_m3_1<>'') or (vla_ed_ppm_1<>'') or (vla_ec_mg_m3_1<>''))  OR ("&monta_condicion_grupo("asoc_vla")&") )" 
				
				case "vl2": 
          campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
          frases = "R40, R45, R49, R40/20, R40/21, R40/22, R40/20/22, R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/21/22, R68/20/21/22"
				  sqls=sqls & "( ((vl.vla_ed_ppm_1 <> '') OR (vl.vla_ed_mg_m3_1 <> '') OR (vl.vla_ec_ppm_1 <> '') OR (vl.vla_ec_mg_m3_1 <> '') OR (vl.vla_ed_ppm_2 <> '') OR (vl.vla_ed_mg_m3_2 <> '') OR (vl.vla_ec_ppm_2 <> '') OR (vl.vla_ec_mg_m3_2 <> '') OR (vl.vla_ed_ppm_3 <> '') OR (vl.vla_ed_mg_m3_3 <> '') OR (vl.vla_ec_ppm_3 <> '') OR (vl.vla_ec_mg_m3_3 <> '') OR (vl.vla_ed_ppm_4 <> '') OR (vl.vla_ed_mg_m3_4 <> '') OR (vl.vla_ec_ppm_4 <> '') OR (vl.vla_ec_mg_m3_4 <> '') OR (vl.vla_ed_ppm_5 <> '') OR (vl.vla_ed_mg_m3_5 <> '') OR (vl.vla_ec_ppm_5 <> '') OR (vl.vla_ec_mg_m3_5 <> '') OR (vl.vla_ed_ppm_6 <> '') OR (vl.vla_ed_mg_m3_6 <> '') OR (vl.vla_ec_ppm_6 <> '') OR (vl.vla_ec_mg_m3_6 <> '')) and ("&monta_condicion(campos, frases)&")  OR ("&monta_condicion_grupo("asoc_vla")&") )" 
				
				case "vl3": 
				  sqls=sqls & "( ((vlb_1<>''))  OR ("&monta_condicion_grupo("asoc_vlb")&") )" 
				
				case "enf": 'la condicion simplemente es que exista la fila, pero ya que estamos, comprobamos que no este vacía
				  sqls=sqls & "( ((sus.id<>'')) OR ("&monta_condicion_grupo("asoc_enfermedades")&") )" 
				
				case "emi": 
				  sqls=sqls & "( ((dn_risc_sustancias_ambiente.emisiones_atmosfera=1)) OR ("&monta_condicion_grupo("asoc_emisiones_atmosfericas")&") )"
				
				case "res": 'la condicion es que exista la fila O QUE que num_rd <>'' / como tenemos full outer, añadimos la condicion de que la tabla de sustancias no este vacia, por si acaso
				  sqls=sqls & " ((sus.num_rd<>'' or sus.num_rd <> '' or dn_risc_sustancias_ambiente.id is not null or dn_risc_sustancias_cancer_otras.id is not null or dn_risc_sustancias_iarc.id is not null or dn_risc_sustancias_neuro_disruptor.id is not null or dn_risc_sustancias_vl.id is not null ) AND sus.id is not null) " 
				
				case "ver": 
				  sqls=sqls & " ((sus.num_rd <> '') OR (sus.frases_r_danesa <> '') OR (iarc.grupo_iarc <> '') OR (otras.categoria_cancer_otras <> '') OR (neuro.nivel_disruptor <> '') OR (ambiente.enlace_tpb <> '') OR (ambiente.directiva_aguas <> '') OR (ambiente.clasif_mma <> '')) "
				
				case "cov": 
				  sqls=sqls & " ((dn_risc_sustancias_ambiente.cov=1)) "
				
				case "lpc": 
				  sqls=sqls & "( ((eper_agua<>'' or eper_aire<>'')) OR ("&monta_condicion_grupo("asoc_eper")&") )" 
				
				case "acm": 
				  sqls=sqls & "( ((seveso<>'')) OR ("&monta_condicion_grupo("asoc_seveso")&") )" 

        case "negra": 'Lista negra
				  sqls=sqls & " ((negra=1))" 
				
			end select
			
		end if
		
		sqls=sqls & " ORDER BY " &ordenacion&  " "
		'response.write sqls
		
			Set objRst = Server.CreateObject("ADODB.Recordset")
			objRst.Open sqls, objConnection2, adOpenStatic, adCmdText
			hr=objRst.recordcount
		
			IF not objRst.eof THEN 
				'arr=objRst.GetString(adClipString, 1, "", ",", "")		
				arrayDatos=objRst.getrows	
				
				for I = 0 to UBound(arrayDatos,2) 
					arr=arr& arrayDatos(0,I) & ","
				next	
				'esta sera la pagina 1
				pag = 1	
			END IF
					
			objRst.Close
			Set objRst=Nothing

	case 2: 'paginando
		
		hr=request("hr")
		pag=request("pag")
		arr=request("arr")
					
	end select 'cual busc
	
	'RESULTADOS DE BUSQUEDA (para busc 1 y busc 2)
	'seleccionamos datos a mostrar de los x registros que toquen
	if hr>0 then
	
			'vemos que registros hay que mostrar
			registroini=(pag*nregs)-nregs
			'response.write "<p>registroini=" &registroini& "</p>"
			
			registrofin=registroini+nregs
			'response.write "<p>registrofin=" &registrofin& "</p>"
			
			if registrofin>=hr-1 then
				registrofin=hr
      			'response.write "<p>registrofin era mayor, ahora=" &registrofin& "</p>"
			end if
			
			registrofin=registrofin-1
			'response.write "<p>registrofin corregido=" &registrofin& "</p>"
				
		arrayx = split(arr, ",")
		
		FOR i=registroini to registrofin
			cadenaids=cadenaids  &arrayx(i)&","		
		NEXT	
		
		'quitamos la ultima coma
		cadenaids= left(cadenaids,len(cadenaids)-1)
		sqlpag="select id, nombre from dn_risc_sustancias as sus WHERE id IN(" &cadenaids& ") ORDER BY " &ordenacion&  " " &sentido
		'response.write sqlpag
		set rstpag=objConnection2.execute(sqlpag)
		if not rstpag.eof then
			'strDBDataTable = rstpag.GetString(adClipString, -1, "</td><td>", "</td></tr>" & vbCrLf & "<tr><td>", "&nbsp;")
			'strDBDataTable= left(strDBDataTable,len(strDBDataTable)-8)
			'tablares= "<table border='1'><tr><td>" &strDBDataTable& "</table>" 
			arrayDatos = rstpag.GetRows		

			'for contadorFilas=registroini to registrofin		
			for contadorFilas=0 to registrofin-registroini
					
					'if contadorfilas>=hr-1 then
							'exit for
					'else
						  'arrayDatos(0,contadorFilas)
						  tablares=tablares & "<tr>" 
						  'tablares=tablares & "<td>" & contadorFilas+1 & "</td>"
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
		
	end if
end if 'busc

if hr=1 then
			response.redirect(unico_enlace)
end if
%>
