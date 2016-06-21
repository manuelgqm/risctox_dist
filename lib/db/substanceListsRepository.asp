<%
function isSubstanceInList(byval lista, byval id_sustancia, connection)

	' Montamos condicion inicial dependiendo de lista, como en buscador publico de risctox pero sin sinónimos
	select case lista
		case "cancer_rd": ' Cancerigeno según RD
      sql_lista = whereClause(sql_lista_cancer_rd) & " OR ("&monta_condicion_grupo("asoc_cancer_rd")&") )"

		case "cancer_danesa": ' Cancerigeno según lista danesa
      sql_lista = sql_lista_cancer_danesa

		case "mutageno_rd": ' Mutágeno según RD
      sql_lista = sql_lista_mutageno_rd


		case "mutageno_danesa": ' Mutágeno según lista danesa
      sql_lista = sql_lista_mutageno_danesa

		case "cancer_iarc": ' Cancerígena según IARC
      sql_lista = whereClause(sql_lista_cancer_iarc) & " OR ("&monta_condicion_grupo("asoc_cancer_iarc")&") )"

		case "cancer_iarc_excepto_grupo_3": ' Cancerígena según IARC, excepto Grupo 3
      sql_lista = sql_lista_cancer_iarc_excepto_grupo_3

		case "cancer_otras": ' Cancerígena según otras fuentes
      sql_lista = whereClause(sql_lista_cancer_otras) & " OR ("&monta_condicion_grupo("asoc_cancer_otras")&") )"

		case "cancer_otras_excepto_grupo_4":
      sql_lista = whereClause(sql_lista_cancer_otras) & " OR ("&monta_condicion_grupo("asoc_cancer_otras")&") ) AND dn_risc_sustancias_cancer_otras.categoria_cancer_otras not like '%G-A4%'"

		case "cancer_mama": ' Cancerígena mama
      sql_lista = whereClause(sql_lista_cancer_mama) & " OR ("&monta_condicion_grupo("asoc_cancer_mama")&") )"

		case "cop": ' COP
      sql_lista = whereClause(sql_lista_cop) & " OR ("&monta_condicion_grupo("asoc_cop")&") )"

		case "salud": ' Efectos para la salud y órganos afectados
      sql_lista = sql_lista_salud

		case "tpr": ' Tóxicos para la reproducción
      sql_lista = whereClause(sql_lista_tpr) & " OR ("&monta_condicion_grupo("asoc_reproduccion")&") )"

		case "tpr_danesa": ' Tóxicos para la reproducción según lista danesa
      sql_lista = sql_lista_tpr_danesa

		case "de": ' Disruptor endocrino
      sql_lista = whereClause(sql_lista_de) & " OR ("&monta_condicion_grupo("asoc_disruptores")&") )"

		case "neurotoxico": ' Neurótoxico (RD o Danesa o por nivel)
'      sql_lista = whereClause(sql_lista_neurotoxico) & " OR ("&monta_condicion_grupo("asoc_neuro_oto")&") )"
      sql_lista = "select distinct sus.id, sus.nombre from dn_risc_sustancias as sus WHERE " & " ("&monta_condicion_grupo("asoc_neuro_oto")&") "

		case "neurotoxico_rd": ' Neurótoxico RD
      sql_lista = sql_lista_neurotoxico_rd

		case "neurotoxico_danesa": ' Neurótoxico Danesa
      sql_lista = sql_lista_neurotoxico_danesa

		case "neurotoxico_nivel": ' Neurótoxico Danesa
      sql_lista = sql_lista_neurotoxico_nivel

		case "sensibilizante": ' Sensibilizante
      sql_lista = sql_lista_sensibilizante

		case "sensibilizante_danesa": ' Sensibilizante según lista danesa
      sql_lista = sql_lista_sensibilizante_danesa

	  case "sensibilizante_reach": ' Sensibilizante según reach
      'sql_lista = sql_lista_sensibilizante_reach
      sql_lista = whereClause(sql_lista_sensibilizante_reach) & " OR ("&monta_condicion_grupo("asoc_alergenos")&") )"

		case "eepp": ' Enfermedades profesionales relacionadas
      sql_lista = sql_lista_eepp

		case "tpb": ' Tóxicas, persistentes y bioacumulativas
      'sql_lista = sql_lista_tpb
      sql_lista = whereClause(sql_lista_tpb) & " OR ("&monta_condicion_grupo("asoc_tpb")&") )"

		case "directiva_aguas": ' Directiva de aguas
      sql_lista = whereClause(sql_lista_directiva_aguas) & " OR ("&monta_condicion_grupo("asoc_directiva_aguas")&") )"
	  'response.write sql_lista

		case "sustancias_prioritarias": '
      sql_lista = sql_lista_sustancia_prioritaria

		case "alemana": ' Alemana de aguas
      'sql_lista = sql_lista_alemana
      sql_lista = whereClause(sql_lista_alemana) & " OR ("&monta_condicion_grupo("asoc_peligrosas_agua_alemania")&") )"

		case "ozono": ' Capa de ozono
      'sql_lista = sql_lista_ozono
      sql_lista = whereClause(sql_lista_ozono) & " OR ("&monta_condicion_grupo("asoc_capa_ozono")&") )"

		case "clima": ' Cambio climático
      sql_lista = sql_lista_clima
      sql_lista = whereClause(sql_lista_clima) & " OR ("&monta_condicion_grupo("asoc_cambio_climatico")&") )"

		case "aire": ' Calidad del aire
      sql_lista = whereClause(sql_lista_aire) & " OR ("&monta_condicion_grupo("asoc_calidad_aire")&") )"

		case "cov": ' COV
      'sql_lista = sql_lista_cov
      sql_lista = whereClause(sql_lista_cov) & " OR ("&monta_condicion_grupo("asoc_cov")&") )"

	  case "suelos": ' Contaminante suelos
      'sql_lista = sql_lista_suelos
      sql_lista = whereClause(sql_lista_suelos) & " OR ("&monta_condicion_grupo("asoc_contaminantes_suelo")&") )"

		case "vertidos": ' Vertidos
      sql_lista = sql_lista_vertidos

		case "lpcic": ' LPCIC
      'sql_lista = whereClause(sql_lista_lpcic) & " OR ("&monta_condicion_grupo("asoc_eper")&") )"
      sql_lista = sql_lista_lpcic

		case "lpcic-agua": ' LPCIC Agua
      'sql_lista = whereClause(sql_lista_lpcic_agua) & " OR ("&monta_condicion_grupo("asoc_eper")&") )"
      sql_lista = sql_lista_lpcic_agua

		case "lpcic-aire": ' LPCIC Aire
      'sql_lista = whereClause(sql_lista_lpcic_aire) & " OR ("&monta_condicion_grupo("asoc_eper")&") )"
      sql_lista = sql_lista_lpcic_aire

	  case "lpcic-suelo": ' LPCIC Aire
      'sql_lista = whereClause(sql_lista_lpcic_aire) & " OR ("&monta_condicion_grupo("asoc_eper")&") )"
      sql_lista = sql_lista_lpcic_suelo

		case "residuos": ' Residuos peligrosos
      sql_lista = sql_lista_residuos

		case "accidentes": ' Accidentes graves
      sql_lista = whereClause(sql_lista_accidentes) & " OR ("&monta_condicion_grupo("asoc_seveso")&") )"

		case "emisiones": ' Emisiones atmosféricas
      sql_lista = whereClause(sql_lista_emisiones) & " OR ("&monta_condicion_grupo("asoc_emisiones_atmosfericas")&") )"

		case "prohibidas": ' Sustancias prohibidas
      		sql_lista = whereClause(sql_lista_prohibidas) & "(sus.id=pro.id_sustancia) OR ("&monta_condicion_grupo("asoc_prohibidas")&"))"

		case "restringidas": ' Sustancias restringidas
      		sql_lista = whereClause(sql_lista_restringidas) & "(sus.id=rest.id_sustancia) OR ("&monta_condicion_grupo("asoc_restringidas")&"))"

	' SPL NUEVAS LISTAS
		case "prohibidas_embarazadas": '
			campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
      		frases="R60, R61"
      		sql_lista = whereClause(sql_lista_prohibidas_embarazadas)
			sql_lista = sql_lista & "(sus.id=pro_emb.id_sustancia) OR  " ' Lista de sustancias prohibidas para embarazadas
			sql_lista = sql_lista & "( " & monta_condicion(campos, frases) & " OR " ' Sustancias con R60 y R61
			sql_lista = sql_lista & " (sus.num_rd='082-001-00-6' ) " ' Sustancias con rd=082-001-00-6
			sql_lista = sql_lista & " OR ("&monta_condicion_grupo("asoc_prohibidas_embarazadas")&" )"

			sql_lista = sql_lista & " OR (sus.num_rd='650-017-00-8' AND sus.num_rd='650-016-00-2')"
			sql_lista = sql_lista & "))"


'response.write sql_lista
		case "prohibidas_lactantes": '
			campos="sus.clasificacion_1, sus.clasificacion_2, sus.clasificacion_3, sus.clasificacion_4, sus.clasificacion_5, sus.clasificacion_6, sus.clasificacion_7, sus.clasificacion_8, sus.clasificacion_9, sus.clasificacion_10, sus.clasificacion_11, sus.clasificacion_12, sus.clasificacion_13, sus.clasificacion_14, sus.clasificacion_15"
     		frases="R64"
      		sql_lista = whereClause(sql_lista_prohibidas_lactantes)
			sql_lista = sql_lista & "(sus.id=pro_lac.id_sustancia) OR  " ' Lista de sustancias prohibidas para lactantes
			sql_lista = sql_lista & "( " & monta_condicion(campos, frases) & " OR " ' Sustancias con R60 y R61
			sql_lista = sql_lista & " (sus.num_rd='082-001-00-6' ) " ' Sustancias con rd=082-001-00-6
			sql_lista = sql_lista & " OR ("&monta_condicion_grupo("asoc_prohibidas_lactantes")&" )"

			sql_lista = sql_lista & " OR (sus.num_rd='650-017-00-8' AND sus.num_rd='650-016-00-2')"
			sql_lista = sql_lista & "))"

		case "candidatas_reach":
			sql_lista = whereClause(sql_lista_candidatas_reach) & "(sus.id=candidatas_reach.id_sustancia)  OR ("&monta_condicion_grupo("asoc_candidatas_reach")&"))"

		case "autorizacion_reach":
			sql_lista = whereClause(sql_lista_autorizacion_reach) & "(sus.id=autorizacion_reach.id_sustancia)  OR ("&monta_condicion_grupo("asoc_autorizacion_reach")&"))"

		case "biocidas_autorizadas":
			sql_lista = whereClause(sql_lista_biocidas_autorizadas) & "(sus.id=biocidas_autorizadas.id_sustancia)  OR ("&monta_condicion_grupo("asoc_biocidas_autorizadas")&"))"

		case "biocidas_prohibidas":
			sql_lista = whereClause(sql_lista_biocidas_prohibidas) & "(sus.id=biocidas_prohibidas.id_sustancia)  OR ("&monta_condicion_grupo("asoc_biocidas_prohibidas")&"))"

		case "pesticidas_autorizadas":
			sql_lista = whereClause(sql_lista_pesticidas_autorizadas) & "(sus.id=pesticidas_autorizadas.id_sustancia)  OR ("&monta_condicion_grupo("asoc_pesticidas_autorizadas")&"))"

		case "pesticidas_prohibidas":
			sql_lista = whereClause(sql_lista_pesticidas_prohibidas) & "(sus.id=pesticidas_prohibidas.id_sustancia)  OR ("&monta_condicion_grupo("asoc_pesticidas_prohibidas")&"))"

		case "corap"
			sql_lista = sql_lista_corap & "(sus.id = sustancias_corap.id_sustancia)"

	end select

	sql_lista = sql_lista & " AND sus.id = " & id_sustancia

	set obj_rst_lista = connection.execute(sql_lista)
	if (obj_rst_lista.eof) then
		esta = false
	else
		esta = true
	end if

	obj_rst_lista.close()
	set obj_rst_lista = nothing

	esta_en_lista = esta
end function

function whereClause(byval cadena)
  ' Añade otro paréntesis al principio del WHERE
  cadena = ucase(cadena)
  whereClause = replace(cadena, "WHERE", "WHERE (")
end function

%>
