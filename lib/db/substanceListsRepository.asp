<%
function obtainFeaturedLists(id_sustancia, connection)
	dim substanceLists
	substanceLists = Array( _
		"cancer_rd", "cancer_danesa", "mutageno_rd", "mutageno_danesa", _
		"cancer_iarc", "cancer_iarc_excepto_grupo_3", "cancer_otras", "cancer_mama", _
		"tpr", "tpr_danesa", "de", "neurotoxico_rd", "neurotoxico_danesa", _
		"neurotoxico_nivel", "neurotoxico" , "sensibilizante", "sensibilizante_danesa", _
		"sensibilizante_reach", "eepp", "tpb", "directiva_aguas", _
		"sustancias_prioritarias", "alemana", "aire", "ozono", "clima", _
		"suelos", "cov", "vertidos", "lpcic", "lpcic-agua", "lpcic-aire", _
		"lpcic-suelo", "residuos", "accidentes", "emisiones", "salud", _
		"prohibidas", "restringidas", "cop", "prohibidas_embarazadas", _
		"prohibidas_lactantes", "candidatas_reach", "autorizacion_reach", _
		"biocidas_autorizadas", "biocidas_prohibidas", "pesticidas_autorizadas", _
		"pesticidas_prohibidas", "corap" _
	)

	dim listsContainingSubstance()
	dim size
	size = 0
	for i = 0 to uBound(substanceLists)
	  listName = substanceLists(i)
	  if isSubstanceInList(listName, id_sustancia, objConnection2) then
	    redim preserve listsContainingSubstance(size)
	    listsContainingSubstance(size) = listName
	    size = size + 1
	  end if
	next
	obtainFeaturedLists = listsContainingSubstance

end function

function isSubstanceInList(byval lista, byval id_sustancia, connection)
	const frasesRCancer = "R40, R45, R49, R40/20, R40/21, R40/22, R40/20/21, R40/20/22, R40/21/22, R40/20/21/22"
	const frasesRMutageno = "R46, R68, R68/20, R68/21, R68/22, R68/20/21, R68/20/22, R68/21/22, R68/20/21/22"
	const frasesRTpr = "R60, R61, R62, R63"
	const frasesRNeurotoxico = "R67"
	dim isSubstanceInListResult
	isSubstanceInListResult = false

	select case lista
		case "cancer_rd": ' Cancerigeno según RD
      campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
      sqlQuery = parentesis_where( _
        "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & monta_condicion(campos, frasesRCancer) _
      ) &_
      " OR (" & monta_condicion_grupo("asoc_cancer_rd") & ") )"

		case "cancer_danesa": ' Cancerigeno según lista danesa
      campos="sus.frases_r_danesa"
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & monta_condicion(campos, frasesRCancer)

		case "mutageno_rd": ' Mutágeno según RD
			campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
			sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & monta_condicion(campos, frasesRMutageno)

		case "mutageno_danesa": ' Mutágeno según lista danesa
		campos="sus.frases_r_danesa"
		sqlQuery="select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & monta_condicion(campos, frasesRMutageno)

		case "cancer_iarc": ' Cancerígena según IARC
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia) WHERE (dn_risc_sustancias_iarc.grupo_iarc<>'')" _
			) & _
			" OR (" & groupByClause("asoc_cancer_iarc") & ") )"

		case "cancer_iarc_excepto_grupo_3": ' Cancerígena según IARC, excepto Grupo 3
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia) WHERE (dn_risc_sustancias_iarc.grupo_iarc<>'' AND dn_risc_sustancias_iarc.grupo_iarc NOT LIKE '%3%')"

		case "cancer_otras": ' Cancerígena según otras fuentes
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia) WHERE (dn_risc_sustancias_cancer_otras.categoria_cancer_otras<>'')" _
			) & _
			" OR ("& groupByClause("asoc_cancer_otras") & ") )"

		case "cancer_otras_excepto_grupo_4":
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia) WHERE (dn_risc_sustancias_cancer_otras.categoria_cancer_otras<>'')" _
			) & _
			" OR (" & groupByClause("asoc_cancer_otras") & ") )" &_
			" AND dn_risc_sustancias_cancer_otras.categoria_cancer_otras not like '%G-A4%'"

		case "cancer_mama": ' Cancerígena mama
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia) WHERE (dn_risc_sustancias_mama_cop.cancer_mama=1)" _
			) & _
			" OR (" & groupByClause("asoc_cancer_mama") & ") )"

		case "cop": ' COP
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sustancias_mama_cop ON (sus.id=dn_risc_sustancias_mama_cop.id_sustancia) WHERE (dn_risc_sustancias_mama_cop.cop<>'')" _
			) & _
			" OR (" & groupByClause("asoc_cop") & ") )"

		case "salud": ' Efectos para la salud y órganos afectados
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_salud AS sal ON (sus.id=sal.id_sustancia) WHERE (sal.cardiocirculatorio=1 OR sal.rinyon=1 OR sal.respiratorio=1 OR sal.reproductivo=1 OR sal.piel_sentidos=1 OR sal.neuro_toxicos=1 OR sal.musculo_esqueletico=1 OR sal.sistema_inmunitario=1 OR sal.higado_gastrointestinal=1 OR sal.sistema_endocrino=1 OR sal.embrion=1 OR sal.cancer=1)"

		case "tpr": ' Tóxicos para la reproducción
			campos = "sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & monta_condicion(campos, frasesRTpr) _
			) & _
			" OR (" & groupByClause("asoc_reproduccion") & ") )"

		case "tpr_danesa": ' Tóxicos para la reproducción según lista danesa
			campos = "sus.frases_r_danesa"
			sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & monta_condicion(campos, frasesRTpr)

		case "de": ' Disruptor endocrino
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia) WHERE (dn_risc_sustancias_neuro_disruptor.nivel_disruptor<>'')" _
			) & _
			" OR (" & groupByClause("asoc_disruptores") & ") )"

		case "neurotoxico": ' Neurótoxico RD, Danesa o nivel
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & " (" & groupByClause("asoc_neuro_oto") & ") "

		case "neurotoxico_rd":
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & monta_condicion(campos, frasesRNeurotoxico)

		case "neurotoxico_danesa":
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & monta_condicion(campos, frasesRNeurotoxico)

		case "neurotoxico_nivel": ' Neurótoxico Danesa
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia) WHERE (dn_risc_sustancias_neuro_disruptor.nivel_neurotoxico<>'')"

		case "sensibilizante": ' Sensibilizante
			const frasesRSensibilizante = "R42, R43, R42/43, R42-43"
			campos = "sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & monta_condicion(campos, frasesRSensibilizante)

		case "sensibilizante_danesa": ' Sensibilizante según lista danesa
			const frasesRSensibilizanteDanesa = "R42, R43, R42/43"
			campos="sus.frases_r_danesa"
    	sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE "&monta_condicion(campos, frasesRSensibilizanteDanesa)

	  case "sensibilizante_reach": ' Sensibilizante según reach
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sensibilizantes_reach AS sen ON (sus.id=sen.id_sustancia)  WHERE (sus.id<>'' AND sen.id_sustancia <> '')" ) _
			 & " OR (" & groupByClause("asoc_alergenos") & ") )"

		case "eepp": ' Enfermedades profesionales relacionadas
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sustancias_por_grupos AS spg ON (sus.id=spg.id_sustancia) LEFT OUTER JOIN dn_risc_grupos_por_enfermedades AS gpe ON (spg.id_grupo=gpe.id_grupo) LEFT OUTER JOIN dn_risc_sustancias_por_enfermedades AS spe ON sus.id = spe.id_sustancia WHERE ((sus.id<>'' AND (spe.id_enfermedad IS NOT NULL) OR (gpe.id_enfermedad IS NOT NULL)))"

		case "tpb": ' Tóxicas, persistentes y bioacumulativas
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.anchor_tpb<>'')" ) _
			& " OR (" & groupByClause("asoc_tpb") & ") )"

		case "directiva_aguas": ' Directiva de aguas
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.directiva_aguas=1)" _
			) & _
			" OR (" & groupByClause("asoc_directiva_aguas") & ") )"

		case "sustancias_prioritarias": '
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (sustancia_prioritaria=1)"

		case "alemana": ' Alemana de aguas
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sinonimos as sin ON (sus.id=sin.id_sustancia) LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.clasif_MMA<>''and DN_RISC_SUSTANCIAS_AMBIENTE.CLASIF_MMA <> 'nwg')" _
				) & _
				" OR (" & groupByClause("asoc_peligrosas_agua_alemania") & ") )"

		case "ozono": ' Capa de ozono
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.dano_ozono=1)" _
			) & _
		 " OR (" & groupByClause("asoc_capa_ozono") & ") )"

		case "clima": ' Cambio climático
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.dano_cambio_clima=1)" _
			) & _
		 " OR (" & groupByClause("asoc_cambio_climatico") & ") )"

		case "aire": ' Calidad del aire
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.dano_calidad_aire=1)" _
			) & _
			" OR (" & groupByClause("asoc_calidad_aire") & ") )"

		case "cov": ' COV
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.cov=1)" _
				) & _
			" OR (" & groupByClause("asoc_cov") & ") )"

	  case "suelos": ' Contaminante suelos
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.toxicidad_suelo=1)" _
			) & _
			" OR (" & groupByClause("asoc_contaminantes_suelo") & ") )"

		case "vertidos": ' Vertidos
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_iarc iarc ON sus.id = iarc.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_cancer_otras otras ON sus.id = otras.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_neuro_disruptor neuro ON sus.id = neuro.id_sustancia LEFT OUTER JOIN dn_risc_sustancias_ambiente ambiente ON sus.id = ambiente.id_sustancia WHERE ((sus.num_rd <> '') OR (sus.frases_r_danesa <> '') OR (iarc.grupo_iarc <> '') OR (otras.categoria_cancer_otras <> '') OR (neuro.nivel_disruptor <> '') OR (ambiente.enlace_tpb <> '') OR (ambiente.directiva_aguas <> '') OR (ambiente.clasif_mma <> ''))"

		case "lpcic": ' LPCIC
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (eper_agua<>'' or eper_aire<>'' or eper_suelo<>'')"

		case "lpcic-agua": ' LPCIC Agua
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (eper_agua<>'')"

		case "lpcic-aire": ' LPCIC Aire
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (eper_aire<>'')"

	  case "lpcic-suelo": ' LPCIC Aire
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (eper_suelo<>'')"

		case "residuos": ' Residuos peligrosos
      sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus FULL OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) FULL OUTER JOIN dn_risc_sustancias_cancer_otras ON (sus.id=dn_risc_sustancias_cancer_otras.id_sustancia) FULL OUTER JOIN dn_risc_sustancias_iarc ON (sus.id=dn_risc_sustancias_iarc.id_sustancia) FULL OUTER JOIN dn_risc_sustancias_neuro_disruptor ON (sus.id=dn_risc_sustancias_neuro_disruptor.id_sustancia) FULL OUTER JOIN dn_risc_sustancias_vl ON (sus.id=dn_risc_sustancias_vl.id_sustancia) WHERE ((sus.num_rd<>'' or sus.frases_r_danesa <> '' or dn_risc_sustancias_ambiente.id is not null or dn_risc_sustancias_cancer_otras.id is not null or dn_risc_sustancias_iarc.id is not null or dn_risc_sustancias_neuro_disruptor.id is not null or dn_risc_sustancias_vl.id is not null ) AND sus.id is not null)"

		case "accidentes": ' Accidentes graves
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (seveso<>'')" _
			) & _
		 " OR (" & groupByClause("asoc_seveso") & ") )"

		case "emisiones": ' Emisiones atmosféricas
      sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_ambiente ON (sus.id=dn_risc_sustancias_ambiente.id_sustancia) WHERE (dn_risc_sustancias_ambiente.emisiones_atmosfera=1)" _
			) & _
		 " OR (" & groupByClause("asoc_emisiones_atmosfericas") & ") )"

		case "prohibidas": ' Sustancias prohibidas
			sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_prohibidas as pro ON (sus.id=pro.id_sustancia) WHERE " _
			) & _
			"(sus.id=pro.id_sustancia) OR (" & groupByClause("asoc_prohibidas") & "))"

		case "restringidas": ' Sustancias restringidas
      		sqlQuery = whereClause( _
						"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN dn_risc_sustancias_restringidas as rest ON (sus.id=rest.id_sustancia) WHERE " _
					) & _
				 "(sus.id=rest.id_sustancia) OR (" & groupByClause("asoc_restringidas") & "))"

	' SPL NUEVAS LISTAS
		case "prohibidas_embarazadas": '
			const frasesREmbarazadas = "R60, R61"
			campos = "sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
  		sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_prohibidas_embarazadas as pro_emb ON (sus.id=pro_emb.id_sustancia) WHERE " _
			) & _
			"(sus.id=pro_emb.id_sustancia) OR  " & _
			"( " & buildCondition(campos, frasesREmbarazadas) & " OR "  & _
			" (sus.num_rd='082-001-00-6' ) " & _
			" OR (" & groupByClause("asoc_prohibidas_embarazadas") & " )" & _
			" OR (sus.num_rd='650-017-00-8' AND sus.num_rd='650-016-00-2')" & _
			"))"

		case "prohibidas_lactantes": '
     	const frasesRLactantes="R64"
			campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
  		sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_prohibidas_lactantes as pro_lac ON (sus.id=pro_lac.id_sustancia) WHERE " _
			) & _
			"(sus.id=pro_lac.id_sustancia) OR  " & _
			"( " & buildCondition(campos, frasesRLactantes) & " OR " & _
			" (sus.num_rd='082-001-00-6' ) " & _
			" OR (" & groupByClause("asoc_prohibidas_lactantes") & " )" & _
			" OR (sus.num_rd='650-017-00-8' AND sus.num_rd='650-016-00-2')" & _
			"))"

		case "candidatas_reach":
			sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_candidatas_reach as candidatas_reach ON (sus.id=candidatas_reach.id_sustancia) WHERE " _
				) & _
			"(sus.id=candidatas_reach.id_sustancia)  OR (" & groupByClause("asoc_candidatas_reach") & "))"

		case "autorizacion_reach":
			sqlQuery = whereClause( _
			"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_autorizacion_reach as autorizacion_reach ON (sus.id=autorizacion_reach.id_sustancia) WHERE " _
			) &_
			"(sus.id=autorizacion_reach.id_sustancia)  OR (" & groupByClause("asoc_autorizacion_reach") & "))"

		case "biocidas_autorizadas":
			sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_biocidas_autorizadas as biocidas_autorizadas ON (sus.id=biocidas_autorizadas.id_sustancia) WHERE " _
			) & _
			"(sus.id=biocidas_autorizadas.id_sustancia)  OR (" & groupByClause("asoc_biocidas_autorizadas") & "))"

		case "biocidas_prohibidas":
			sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_biocidas_prohibidas as biocidas_prohibidas ON (sus.id=biocidas_prohibidas.id_sustancia) WHERE " _
			) & _
			"(sus.id=biocidas_prohibidas.id_sustancia)  OR (" & groupByClause("asoc_biocidas_prohibidas") & "))"

		case "pesticidas_autorizadas":
			sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_pesticidas_autorizadas as pesticidas_autorizadas ON (sus.id=pesticidas_autorizadas.id_sustancia) WHERE " _
			) & _
			"(sus.id=pesticidas_autorizadas.id_sustancia)  OR (" & groupByClause("asoc_pesticidas_autorizadas") & "))"

		case "pesticidas_prohibidas":
			sqlQuery = whereClause( _
				"select distinct sus.id, sus.nombre from dn_risc_sustancias as sus LEFT OUTER JOIN spl_risc_sustancias_pesticidas_prohibidas as pesticidas_prohibidas ON (sus.id=pesticidas_prohibidas.id_sustancia) WHERE " _
			) & _
			"(sus.id=pesticidas_prohibidas.id_sustancia)  OR (" & groupByClause("asoc_pesticidas_prohibidas") & "))"

		case "corap"
			sqlQuery = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus left outer join ist_risc_sustancias_corap as sustancias_corap ON (sus.id = sustancias_corap.id_sustancia) WHERE (sus.id = sustancias_corap.id_sustancia)"

	end select

	sqlQuery = sqlQuery & " AND sus.id = " & id_sustancia

	dim recordsetLista
	set recordsetLista = connection.execute(sqlQuery)

	if not recordsetLista.eof then isSubstanceInListResult = true

	recordsetLista.close()
	set recordsetLista = nothing

	isSubstanceInList = isSubstanceInListResult
end function

' PRIVATE
function whereClause(byval cadena)
  ' Añade otro paréntesis al principio del WHERE
  cadena = ucase(cadena)
  whereClause = replace(cadena, "WHERE", "WHERE (")
end function

function groupByClause(check_lista)
  ' Devuelve una cadena para incluir sustancias asociadas a través de grupo en las listas de risctox
  ' Se le debe indicar por parámetro el nombre del checkbox correspondiente a la lista en el formulario de grupo
  ' de la herramienta (ejemplo, "asoc_cop")

   groupByClause = "sus.id IN (SELECT DISTINCT spg.id_sustancia FROM dn_risc_grupos AS g INNER JOIN dn_risc_sustancias_por_grupos AS spg ON spg.id_grupo = g.id WHERE g." & check_lista & "=1)"

end function

function buildCondition(byval campos, byval frases)
  ' Helper para montar la parte de SQL donde se buscan frases R en los campos clasificacion_xx y/o frases_r_danesa,
  ' indicando en qué campos buscar (separados por comas) y qué frases (tambien separados por comas)

  ' Ejemplo:
  ' buildCondition("sus.clasificacion_1, sus_clasificacion_2, sus_clasificacion_3", "R42, R43, R42/43") devuelve:
  ' "((sus.clasificacion_1 LIKE '%R42') OR (sus.clasificacion_1 LIKE '%R42;%') OR (sus.clasificacion_1 LIKE '%R43') OR (sus.clasificacion_1 LIKE '%R43;%') OR (sus.clasificacion_1 LIKE '%R42/43') OR (sus.clasificacion_1 LIKE '%R42/43;%') OR (sus.clasificacion_2 LIKE '%R42') OR (sus.clasificacion_2 LIKE '%R42;%') OR (sus.clasificacion_2 LIKE '%R43') OR (sus.clasificacion_2 LIKE '%R43;%') OR (sus.clasificacion_2 LIKE '%R42/43') OR (sus.clasificacion_2 LIKE '%R42/43;%') OR (sus.clasificacion_3 LIKE '%R42') OR (sus.clasificacion_3 LIKE '%R42;%') OR (sus.clasificacion_3 LIKE '%R43') OR (sus.clasificacion_3 LIKE '%R43;%') OR (sus.clasificacion_3 LIKE '%R42/43') OR (sus.clasificacion_3 LIKE '%R42/43;%'))"

  ' Creamos array campos y de frases con split
  array_campos = split(campos, ",")
  array_frases = split(frases, ",")

  ' Bucleamos para ir montando la condición
  condicion = ""
  for c=0 to ubound(array_campos)
    ' Para cada campo montamos la variante de frase limpia o acabada en punto y coma
    campo = trim(array_campos(c))

    'Bucleamos para cada frase
    for f=0 to ubound(array_frases)
      frase = trim(array_frases(f))
      if (condicion <> "") then
        condicion = condicion&" OR "
      end if
      ' Buscamos al inicio del campo, o separado por ; o separado por espacio (lista danesa)
      condicion = condicion&"("&campo&" LIKE '%"&frase&"') OR ("&campo&" LIKE '%"&frase&";%') OR ("&campo&" LIKE '%"&frase&" %')"
    next

  next

  buildCondition = "("&condicion&")"

end function
%>
