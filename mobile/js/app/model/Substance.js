define(['Server'], function(Server){
	return function(id){
		var defaults = 
			{ id: id
			, nombre: null
			, sinonimos: null
			, num_cas: null
			, num_ce_einecs: null
			, num_ce_elincs: null
			, num_rd: null
			, nums_icsc: null
			, pictogramasRd1272: null
			, frasesH: null
			, notas_rd1272: null
			, concentracionEtiquetadoRd1272: null
			, valoresLimiteAmbiental: [ 
				{ estado: null
				, ec_ppm: null
				, ec_mg_m3: null
				, ed_ppm: null
				, ed_mg_m3: null
				, notas: [
					{ description: null
					, id: null
					, key: null
					} ]
				} ]
			, valoresLimiteBiologico: [
				{ indicador: null
				, momento: null
				, valor: null
				, notas: [
					{ description : null
					, id: null
					, key: null
					}]
				} ]
			, grupos: null
			, aplicaciones: null
			, featuredLists: null
			, frasesR: null
			, listaNegraClassifications: null
			, explosiva: null
			, clasificacionesRd363: null
			, concentracionEtiquetadoRd363: null
			, pictogramasRd363: null
			, frasesRDanesa: null
			, frasesS: null
			, notas_rd363: null
			, descriptions: 
				{ nombre: 'Se ha asignado el nombre en español que aparece en la base de datos EINECS del Bureau Europeo de Sustancias Químicas'
				, sinonimos: 'Se han incluido otros nombres encontrados en la normativa o en las bases de datos utilizadas para construir la RISCTOX'
				, num_cas: 'Número asignado por el Chemical Abstract Service. Es el sistema de identificación más utilizado a nivel internacional'
				, num_ce_einecs: 'Catálogo europeo de sustancias químicas comercializadas. Dicho inventario establece la lista definitiva de todas las sustancias que en principio se encontraban en el mercado comunitario al 18 de septiembre de 1981.'
				, num_rd: 'El número índice es el número de identificación asignado a la sustancia en el anexo VI del Reglamento 1272/2008 (conocido como CLP) de clasificación, etiquetado y envasado de sustancias y mezclas peligrosas'
				, listaNegra: 'La lista negra incluye aquellas sustancias cuyos posibles daños a la salud y al medio ambiente son tan importantes que debemos evitar su uso o presencia en los lugares de trabajo y su vertido al medio ambiente. Estas sustancias, cuya eliminación será prioritaria, son: Cancerígenas, Mutágenas, Tóxicas para la Reproducción, Disruptores Endocrinos, Sensibilizantes, Neurotóxicos, Tóxicas, Persistentes y Bioacumulativas, muy persistent y muy bioacumulativas, Que pueden provocar a largo plazo efectos negativos en el medio ambiente acuático'
				, clp: 'Incluye las indicaciones de peligro (frases H) asignadas a las sustancias incluidas en el Anexo VI del Reglamento 1272/2008 (CLP). Los grupos de sustancias incluidos en el Reglamento 1272/2008, Ej. compuestos de berilio, se han desglosado en la relación de sustancias que forman este grupo y se les han asignado las frases H y los pictogramas y palabras de advertencia pertenecientes al grupo.	Las sustancias que pertenecen a 2 grupos, Ej. cromato de mercurio, se han clasificado asignando las frases de mayor riesgo de cada grupo o sus combinaciones, según los criterios del Reglamento CLP.'
				, vlaEc: 'Es el valor de referencia para la Exposición de Corta Duración (EC), que es la concentración media del agente químico en la zona de respiración del trabajador, medida o calculada para cualquier período de 15 minutos a lo largo de la jornada laboral, excepto para aquellos agentes químicos para los que se especifique un período de referencia inferior, en la lista de Valores Límite.Lo habitual es determinar las EC de interés, es decir, las del período o períodos de máxima exposición, tomando muestras de 15 minutos de duración en cada uno de ellos. El VLA-EC no debe ser superado por ninguna EC a lo largo de la jornada laboral.Para aquellos agentes químicos que tienen efectos agudos reconocidos pero cuyos principales efectos tóxicos son de naturaleza crónica, el VLA-EC constituye un complemento del VLA-ED y, por tanto, la exposición a estos agentes habrá de valorarse en relación con ambos límites.En cambio, a los agentes químicos de efectos principalmente agudos como, por ejemplo, los gases irritantes, sólo se les asigna para su valoración un VLA-EC.'
				, vlaEd: ''
				, vlaEstado: ''
			}
		}
		Object.assign(this, defaults);
	}
});