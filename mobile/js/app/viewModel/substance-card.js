define(function(){
	return {
		name: 'formaldehído',
		synonyms: 'formaldehído (concentracion 90 por 100), formaldehído . . . %, formol',
		cas: '50-00-0',
		einecs: '200-001-8',
		indice: '605-001-00-5',
		listaNegra: ['cancerígena', 'mutágena', 'neurotóxica', 'sensibilizante'],
		pictograms: [{
			name: 'Toxicidad crónica',
			iconUrl: '../imagenes/pictogramas/pictograma_sensibilizacion_respiratoria.png'
		}, {
			name: 'Toxicidad aguda (oral, cutánea, por inhalación)',
			iconUrl: '../imagenes/pictogramas/pictograma_toxicidad_aguda.png'
		}, {
			name: 'Corrosivo',
			iconUrl: '../imagenes/pictogramas/pictograma_corrosivo_metales.png'
		}, {
			name: 'Peligro',
			iconUrl: ''
		}],
		descriptions: {
			synonyms: 'Se han incluido otros nombres encontrados en la normativa o en las bases de datos utilizadas para construir la RISCTOX',
			cas: 'Número asignado por el Chemical Abstract Service. Es el sistema de identificación más utilizado a nivel internacional',
			einecs: 'Catálogo europeo de sustancias químicas comercializadas. Dicho inventario establece la lista definitiva de todas las sustancias que en principio se encontraban en el mercado comunitario al 18 de septiembre de 1981.',
			indice: 'El número índice es el número de identificación asignado a la sustancia en el anexo VI del Reglamento 1272/2008 (conocido como CLP) de clasificación, etiquetado y envasado de sustancias y mezclas peligrosas',
			listaNegra: 'La lista negra incluye aquellas sustancias cuyos posibles daños a la salud y al medio ambiente son tan importantes que debemos evitar su uso o presencia en los lugares de trabajo y su vertido al medio ambiente. Estas sustancias, cuya eliminación será prioritaria, son: Cancerígenas, Mutágenas, Tóxicas para la Reproducción, Disruptores Endocrinos, Sensibilizantes, Neurotóxicos, Tóxicas, Persistentes y Bioacumulativas, muy persistent y muy bioacumulativas, Que pueden provocar a largo plazo efectos negativos en el medio ambiente acuático',
			clp: 'Incluye las indicaciones de peligro (frases H) asignadas a las sustancias incluidas en el Anexo VI del Reglamento 1272/2008 (CLP). Los grupos de sustancias incluidos en el Reglamento 1272/2008, Ej. compuestos de berilio, se han desglosado en la relación de sustancias que forman este grupo y se les han asignado las frases H y los pictogramas y palabras de advertencia pertenecientes al grupo.	Las sustancias que pertenecen a 2 grupos, Ej. cromato de mercurio, se han clasificado asignando las frases de mayor riesgo de cada grupo o sus combinaciones, según los criterios del Reglamento CLP.'
		}
	}
})