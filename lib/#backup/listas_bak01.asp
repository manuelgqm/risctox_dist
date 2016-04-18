<%

function get_string_tablas( opcion )
	dim sqls
	select case( opcion )
		case "1":
		' Buscador de alternativas. Muestra las sustancias asociadas como toxicas a un uso para el que hay alternativas no toxicas, y también sustancias asociadas a ficheros de alternativas
		sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_por_usos ON (sus.id=dn_risc_sustancias_por_usos.id_sustancia)"
		sqls = sqls & " LEFT OUTER JOIN dn_alter_ficheros_por_sustancias ON (sus.id=dn_alter_ficheros_por_sustancias.id_sustancia)"
		sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_por_grupos ON (sus.id=dn_risc_sustancias_por_grupos.id_sustancia)"
		sqls = sqls & " LEFT OUTER JOIN dn_risc_grupos_por_usos ON (dn_risc_sustancias_por_grupos.id_grupo=dn_risc_grupos_por_usos.id_grupo)"
		sqls = sqls & " LEFT OUTER JOIN dn_alter_ficheros_por_grupos ON (dn_risc_sustancias_por_grupos.id_grupo=dn_alter_ficheros_por_grupos.id_grupo)"

		case "cym": 'cancerigenos y mutagenos segun RD 363/1995
			'no unimos a mas tablas

		case "cym2": 'cancerigenos y mutagenos segun IARC
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia)"

		case "cym3": 'cancerigenos y mutagenos segun otras
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia)"

		case "mama": 'cáncer de mama
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia)"

		case "cop": 'cop
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia)"

		case "tpr": 	'Tóxicos para la reproducción
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_iarc iarc ON sus.id = iarc.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_cancer_otras otras ON sus.id = otras.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor neuro ON sus.id = neuro.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_ambiente ambiente ON sus.id = ambiente.id_sustancia "

		' ***** NUEVAS LISTAS SPL
		case "pro": 	'Prohibidas
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_prohibidas pro ON (sus.id = pro.id_sustancia) "

		case "rest": 	'Restringidas
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_restringidas rest ON (sus.id = rest.id_sustancia) "

		case "pro_emb": 	'Prohibidas para embarazadas
			sqls = sqls & " LEFT OUTER JOIN spl_risc_sustancias_prohibidas_embarazadas pro_emb ON (sus.id = pro_emb.id_sustancia) "

		case "pro_lac": 	'Prohibidas para lactantes
			sqls = sqls & " LEFT OUTER JOIN spl_risc_sustancias_prohibidas_lactantes pro_lac ON (sus.id = pro_lac.id_sustancia) "

		case "candidatas_reach": 	'candidatas reach
			sqls = sqls & " LEFT OUTER JOIN spl_risc_sustancias_candidatas_reach candidatas_reach ON (sus.id = candidatas_reach.id_sustancia) "

		case "autorizacion_reach": 	'autorizaciÃ³n reach
			sqls = sqls & " LEFT OUTER JOIN spl_risc_sustancias_autorizacion_reach autorizacion_reach ON (sus.id = autorizacion_reach.id_sustancia) "

		case "biocidas_autorizadas": 	'biocidas_autorizadas
			sqls = sqls & " LEFT OUTER JOIN spl_risc_sustancias_biocidas_autorizadas biocidas_autorizadas ON (sus.id = biocidas_autorizadas.id_sustancia) "

		case "biocidas_prohibidas": 	'biocidas_prohibidas
			sqls = sqls & " LEFT OUTER JOIN spl_risc_sustancias_biocidas_prohibidas biocidas_prohibidas ON (sus.id = biocidas_prohibidas.id_sustancia) "

		case "pesticidas_autorizadas": 	'pesticidas_autorizadas
			sqls = sqls & " LEFT OUTER JOIN spl_risc_sustancias_pesticidas_autorizadas pesticidas_autorizadas ON (sus.id = pesticidas_autorizadas.id_sustancia) "

		case "pesticidas_prohibidas": 	'pesticidas_prohibidas
			sqls = sqls & " LEFT OUTER JOIN spl_risc_sustancias_pesticidas_prohibidas pesticidas_prohibidas ON (sus.id = pesticidas_prohibidas.id_sustancia) "

		' ***** FIN NUEVAS LISTAS SPL

		case "dis": 'disruptor endocrino
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia)"

		case "neu": 'neurotoxico
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia)"

		'Sergio
		case "oto": 'neurotoxico
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia)"

		case "sen": ' sensibilizante
			'no hace falta mas tablas que la principal

		case "senreach": 'sensibilizantes reach
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sensibilizantes_reach ON (sus.id=dn_risc_sensibilizantes_reach.id_sustancia)"

		case "pyb": 'Sustancias tóxicas, persistentes y bioacumulativas
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "tac": 'Sustancias de toxicidad acuática según Directiva de aguas
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "tac2": 'Sustancias peligrosas agua Alemania
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "dat": 'Sustancias de daño a la capa de ozono
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "dat2": 'Sustancias cambio climático
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "dat3": 'Sustancias calidad aire
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "vl1": 'Límites de exposición profesional: Valores Límite Ambientales
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia)"

		case "vl2": '	Valores Límite Ambientales Cancerígenos
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_vl as vl ON (sus.id=vl.id_sustancia)"

		case "vl3": 'Límites de exposición profesional: Valores Límite Biológicos
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia)"

		case "enf": 'Enfermedades profesionales (borrador)
			sqls = sqls & " INNER JOIN dn_risc_sustancias_por_grupos ON (sus.id=dn_risc_sustancias_por_grupos.id_sustancia) INNER JOIN dn_risc_grupos_por_enfermedades ON (dn_risc_sustancias_por_grupos.id_grupo=dn_risc_grupos_por_enfermedades.id_grupo)"

		case "res": 'residuos
			sqls = sqls & " FULL OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"
			sqls = sqls & " FULL OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia)"
			sqls = sqls & " FULL OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia)"
			sqls = sqls & " FULL OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia)"
			sqls = sqls & " FULL OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia)"

		case "ver": 'vertidos
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_iarc iarc ON sus.id = iarc.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_cancer_otras otras ON sus.id = otras.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor neuro ON sus.id = neuro.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_ambiente ambiente ON sus.id = ambiente.id_sustancia "

		case "emi": 'emisiones
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "cov": 'Compuestos orgánicos volátiles (COV)
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		'Sergio
		case "lpc": 'Sustancias (LPCIC)
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "ep1":
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "ep2":
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "ep3":
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "acm": 'Sustancias que pueden provocar accidentes graves
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "cos":
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia)"

		case "anexo_reach"
			sqls = sqls & " LEFT OUTER JOIN dn_risc_sustancias_por_usos ON (sus.id=dn_risc_sustancias_por_usos.id_sustancia)"

		case "negra": ' Lista negra
			'Sergio
			sqls = sqls & " FULL OUTER JOIN dn_risc_sustancias_iarc as iarc ON (sus.id=iarc.id_sustancia)"
			sqls = sqls & " FULL OUTER JOIN dn_risc_sustancias_cancer_otras as caoc ON (sus.id=caoc.id_sustancia)"
			sqls = sqls & " FULL OUTER JOIN dn_risc_sustancias_neuro_disruptor as neuro ON (sus.id=neuro.id_sustancia)"
	end select
	
	get_string_tablas = sqls
	
end function

function get_string_codicion( opcion )
	dim sqls, campos, frases
	select case( opcion )
		case "1": 'la sustancia (o el grupo al que pertenece) debe tener un uso toxico, o debe existir una alternativa
			sqls = sqls & " (dn_risc_sustancias_por_usos.toxico=1 OR dn_alter_ficheros_por_sustancias.id_fichero is not null OR dn_risc_grupos_por_usos.toxico=1 OR dn_alter_ficheros_por_grupos.id_fichero is not null) "

		case "cym":
			campos = "sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
			frases="R40, R45, R49, R40/20, R40/21, R40/22, R40/20/21, R40/20/22, R40/21/22, R40/20/21/22, R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/20/22, R68/21/22, R68/20/21/22"

			sqls = sqls & "( " & monta_condicion(campos, frases) & " OR (" & monta_condicion_grupo("asoc_cancer_rd") & ") )"

		case "cym2": 'la condicion simplemente es que exista la fila, pero ya que estamos, comprobamos que no este vacía
			sqls = sqls & "( dn_risc_sustancias_iarc.grupo_iarc<>'GRUPO 4' and ((dn_risc_sustancias_iarc.grupo_iarc<>'')) OR (" & monta_condicion_grupo("asoc_cancer_iarc") & ") )"

		case "cym3": 'la condicion simplemente es que exista la fila, pero ya que estamos, comprobamos que no este vacï¿½a
			sqls = sqls & " (not(dn_risc_sustancias_cancer_otras.fuente like '%ACGIH%' and (dn_risc_sustancias_cancer_otras.categoria_cancer_otras like '%G-A5%' or dn_risc_sustancias_cancer_otras.categoria_cancer_otras like '%G-A4%' ) ) or (dn_risc_sustancias_cancer_otras.categoria_cancer_otras is null))  and ( ((dn_risc_sustancias_cancer_otras.categoria_cancer_otras<>'')) OR ("&monta_condicion_grupo("asoc_cancer_otras")&") )"

		case "mama": 'que cancer_mama sea 1
			sqls = sqls & "( ((dn_risc_sustancias_mama_cop.cancer_mama=1)) OR ("&monta_condicion_grupo("asoc_cancer_mama")&") )"

		case "cop": 'que cop no sea vacï¿½o
			sqls = sqls & "( ((dn_risc_sustancias_mama_cop.cop <> '')) OR ("&monta_condicion_grupo("asoc_cop")&") )"

		case "tpr":
			' Buscando las frases R: TPR R60, R61, R62, R63, en las columnas CLASIFICACION_1, hasta CLASIFICACION_6 del RD 363/1995.

			campos = "sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
			frases = "R60, R61, R62, R63"

			sqls = sqls & "( " & monta_condicion(campos, frases) & " OR ("&monta_condicion_grupo("asoc_reproduccion")&") )"

		
		' ***** NUEVAS LISTAS SPL
		case "pro":
			sqls = sqls & "(sus.id=pro.id_sustancia) OR ("&monta_condicion_grupo("asoc_prohibidas")&")"

		case "rest":
			sqls = sqls & "(sus.id=rest.id_sustancia) OR ("&monta_condicion_grupo("asoc_restringidas")&")"


		case "pro_emb":
			' Buscando las frases R: R60, R61 en las columnas CLASIFICACION_1, hasta CLASIFICACION_15 por el Real Decreto 363/1995 (
			campos = "sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
			frases = "R60, R61"

			sqls = sqls & "(sus.id=pro_emb.id_sustancia) " ' Lista de sustancias prohibidas para embarazadas
			sqls = sqls & " OR ( " & monta_condicion(campos, frases)  ' Sustancias con R60 y R61
			sqls = sqls & " OR (sus.num_rd='082-001-00-6' ) " ' Sustancias con rd=082-001-00-6
			sqls = sqls & " OR ("&monta_condicion_grupo("asoc_prohibidas_embarazadas")&" )"
			sqls = sqls & " OR (sus.num_rd='650-017-00-8' AND sus.num_rd='650-016-00-2')"
			sqls = sqls & ")"

		case "pro_lac":
			' Buscando las frases R: R60, R61 en las columnas CLASIFICACION_1, hasta CLASIFICACION_15 por el Real Decreto 363/1995 (
			campos = "sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
			frases = "R64"

			sqls = sqls & "(sus.id=pro_lac.id_sustancia) " ' Lista de sustancias prohibidas para lactantes
			sqls = sqls & " OR ( " & monta_condicion(campos, frases) ' Sustancias con R64
			sqls = sqls & " OR (sus.num_rd='082-001-00-6' ) " ' Sustancias con rd=082-001-00-6
			sqls = sqls & " OR ("&monta_condicion_grupo("asoc_prohibidas_lactantes")&" )"
			sqls = sqls & " OR (sus.num_rd='650-017-00-8' AND sus.num_rd='650-016-00-2')"
			sqls = sqls & ")"

		case "candidatas_reach":
			sqls = sqls & "(sus.id=candidatas_reach.id_sustancia) OR ("&monta_condicion_grupo("asoc_candidatas_reach")&")"

		case "autorizacion_reach":
			sqls = sqls & "(sus.id=autorizacion_reach.id_sustancia) OR ("&monta_condicion_grupo("asoc_autorizacion_reach")&")"

		case "biocidas_autorizadas":
			sqls = sqls & "(sus.id=biocidas_autorizadas.id_sustancia) OR ("&monta_condicion_grupo("asoc_biocidas_autorizadas")&")"

		case "biocidas_prohibidas":
			sqls = sqls & "(sus.id=biocidas_prohibidas.id_sustancia) OR ("&monta_condicion_grupo("asoc_biocidas_prohibidas")&")"

		case "pesticidas_autorizadas":
			sqls = sqls & "(sus.id=pesticidas_autorizadas.id_sustancia) OR ("&monta_condicion_grupo("asoc_pesticidas_autorizadas")&")"

		case "pesticidas_prohibidas":
			sqls = sqls & "(sus.id=pesticidas_prohibidas.id_sustancia) OR ("&monta_condicion_grupo("asoc_pesticidas_prohibidas")&")"
		' ***** FIN NUEVAS LISTAS SPL

		case "dis": 'nivel_disruptor no esta vacio
			sqls = sqls & "( ((dn_risc_sustancias_neuro_disruptor.nivel_disruptor<>'')) OR ("&monta_condicion_grupo("asoc_disruptores")&") )"

		case "neu": 'en campos adicionales nivel_neurotoxico no esta vacio
			sqls = sqls & "( " & sql_lista_neurotoxico & " OR ("&monta_condicion_grupo("asoc_neuro_oto")&") )"

		'Sergio
		case "oto":
			sqls = sqls & " dn_risc_sustancias_neuro_disruptor.efecto_neurotoxico='OTOTÓXICO'" 

		case "sen": 'determinadas frases R en clasif_1 a clasif_15 y frases_r_danesa
			campos = "sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15, sus.frases_r_danesa"
			frases = "R42, R43, R42/43, R42-43"
			sqls = sqls & monta_condicion(campos, frases)
			
		case "senreach":
			sqls = sqls & " ((dn_risc_sensibilizantes_reach.id_sustancia<>'')) OR ("&monta_condicion_grupo("asoc_alergenos")&")"
			
		case "pyb":
			sqls = sqls & " ((dn_risc_sustancias_ambiente.anchor_tpb<>''))  OR ("&monta_condicion_grupo("asoc_tpb")&")"

		case "tac":
			sqls = sqls & "( ((dn_risc_sustancias_ambiente.directiva_aguas=1)) OR ("&monta_condicion_grupo("asoc_directiva_aguas")&") )"

		case "tac2":
			sqls = sqls & " ((dn_risc_sustancias_ambiente.clasif_MMA<>'' and dn_risc_sustancias_ambiente.clasif_MMA<>'nwg')) OR ("&monta_condicion_grupo("asoc_peligrosas_agua_alemania")&")"

		case "dat":
			sqls = sqls & " ((dn_risc_sustancias_ambiente.dano_ozono=1)) OR ("&monta_condicion_grupo("asoc_capa_ozono")&")"

		case "dat2":
			sqls = sqls & " ((dn_risc_sustancias_ambiente.dano_cambio_clima=1)) OR ("&monta_condicion_grupo("asoc_cambio_climatico")&")"

		case "dat3":
			sqls = sqls & "( ((dn_risc_sustancias_ambiente.dano_calidad_aire=1)) OR ("&monta_condicion_grupo("asoc_calidad_aire")&") )"

		case "vl1":
			sqls = sqls & "( ((vla_ed_ppm_1<>'') or (vla_ed_mg_m3_1<>'') or (vla_ed_ppm_1<>'') or (vla_ec_mg_m3_1<>''))  OR ("&monta_condicion_grupo("asoc_vla")&") )"

		case "vl2":
			campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
			frases = "R40, R45, R49, R40/20, R40/21, R40/22, R40/20/22, R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/21/22, R68/20/21/22"
			
			sqls = sqls & "( ((vl.vla_ed_ppm_1 <> '') OR (vl.vla_ed_mg_m3_1 <> '') OR (vl.vla_ec_ppm_1 <> '') OR (vl.vla_ec_mg_m3_1 <> '') OR (vl.vla_ed_ppm_2 <> '') OR (vl.vla_ed_mg_m3_2 <> '') OR (vl.vla_ec_ppm_2 <> '') OR (vl.vla_ec_mg_m3_2 <> '') OR (vl.vla_ed_ppm_3 <> '') OR (vl.vla_ed_mg_m3_3 <> '') OR (vl.vla_ec_ppm_3 <> '') OR (vl.vla_ec_mg_m3_3 <> '') OR (vl.vla_ed_ppm_4 <> '') OR (vl.vla_ed_mg_m3_4 <> '') OR (vl.vla_ec_ppm_4 <> '') OR (vl.vla_ec_mg_m3_4 <> '') OR (vl.vla_ed_ppm_5 <> '') OR (vl.vla_ed_mg_m3_5 <> '') OR (vl.vla_ec_ppm_5 <> '') OR (vl.vla_ec_mg_m3_5 <> '') OR (vl.vla_ed_ppm_6 <> '') OR (vl.vla_ed_mg_m3_6 <> '') OR (vl.vla_ec_ppm_6 <> '') OR (vl.vla_ec_mg_m3_6 <> '')) and ("&monta_condicion(campos, frases)&")  OR ("&monta_condicion_grupo("asoc_vla")&") )"

		case "vl3":
			sqls = sqls & "( ((vlb_1<>''))  OR ("&monta_condicion_grupo("asoc_vlb")&") )"

		case "enf": 'la condicion simplemente es que exista la fila, pero ya que estamos, comprobamos que no este vacía
			sqls = sqls & "( ((sus.id<>'')) OR ("&monta_condicion_grupo("asoc_enfermedades")&") )"

		case "emi":
			sqls = sqls & "( ((dn_risc_sustancias_ambiente.emisiones_atmosfera=1)) OR ("&monta_condicion_grupo("asoc_emisiones_atmosfericas")&") )"

		case "res": 'la condicion es que exista la fila O QUE que num_rd <>'' / como tenemos full outer, añadimos la condicion de que la tabla de sustancias no este vacia, por si acaso
			sqls = sqls & " ((sus.num_rd<>'' or sus.num_rd <> '' or dn_risc_sustancias_ambiente.id is not null or dn_risc_sustancias_cancer_otras.id is not null or dn_risc_sustancias_iarc.id is not null or dn_risc_sustancias_neuro_disruptor.id is not null or dn_risc_sustancias_vl.id is not null ) AND sus.id is not null) "

		case "ver":
			sqls = sqls & " ((sus.num_rd <> '') OR (sus.frases_r_danesa <> '') OR (iarc.grupo_iarc <> '') OR (otras.categoria_cancer_otras <> '') OR (neuro.nivel_disruptor <> '') OR (ambiente.enlace_tpb <> '') OR (ambiente.directiva_aguas <> '') OR (ambiente.clasif_mma <> '')) "

		case "cov":
			sqls = sqls & " ((dn_risc_sustancias_ambiente.cov=1)) OR ("&monta_condicion_grupo("asoc_cov")&")"

		case "lpc":
			sqls = sqls & "( ((eper_agua<>'' or eper_aire<>'' or eper_suelo<>'')) OR ("&monta_condicion_grupo("asoc_eper")&") )"

		case "ep1":
			sqls = sqls & "(eper_agua<>'')"

		case "ep2":
			sqls = sqls & "(eper_aire<>'')"

		case "ep3":
			sqls = sqls & "(eper_suelo<>'')"

		case "mpmb":
			sqls = sqls & "(sus.num_cas='87-68-3' or sus.num_cas='133-49-3' or sus.num_cas='75-74-1') OR ("&monta_condicion_grupo("asoc_mpmb")&")"

		case "acm":
			sqls = sqls & "( ((seveso<>'')) OR ("&monta_condicion_grupo("asoc_seveso")&") )"

		case "cos":
			sqls = sqls & "( dn_risc_sustancias_ambiente.toxicidad_suelo=1 ) OR ("&monta_condicion_grupo("asoc_contaminantes_suelo")&")"

		case "anexo_reach":
			sqls = sqls & " (dn_risc_sustancias_por_usos.anexo_reach=1)"

		case "negra": 'Lista negra

			campos = "sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
			frases = "R33, R53, R58, R50-53, R51-53, R52-53"
			
			sqls = sqls & "(" & monta_condicion(campos, frases)
			sqls = sqls & " or ((negra=1)"
			sqls = sqls & " and (iarc.grupo_iarc<>'GRUPO 3' or iarc.grupo_iarc is null) and (iarc.grupo_iarc<>'GRUPO 4' or iarc.grupo_iarc is null)"
			sqls = sqls & " and (not(caoc.fuente like '%ACGIH%' and (caoc.categoria_cancer_otras like '%G-A5%' or caoc.categoria_cancer_otras like '%G-A4%')) or (caoc.categoria_cancer_otras is null))"
			sqls = sqls & " and not (sus.clasificacion_1 like '%67%' or sus.clasificacion_2 like '%67%' or sus.clasificacion_3 like '%67%' or sus.clasificacion_4 like '%67%' or sus.clasificacion_5 like '%67%' or sus.clasificacion_6 like '%67%' or sus.clasificacion_7 like '%67%' or sus.clasificacion_8 like '%67%' or sus.clasificacion_9 like '%67%' or sus.clasificacion_10 like '%67%' or sus.clasificacion_11 like '%67%' or sus.clasificacion_12 like '%67%' or sus.clasificacion_13 like '%67%' or sus.clasificacion_14 like '%67%' or sus.clasificacion_15 like '%67%')))"

	end select
	get_string_codicion = sqls
	
end function
%>