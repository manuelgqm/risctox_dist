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
		frasesH: [
			{value: 'H350', description: 'Puede provocar cáncer'}, 
			{value: 'H341', description: 'Se sospecha que provoca defectos genéticos'}, 
			{value: 'H301', description: 'Tóxico en caso de ingestión'}, 
			{value: 'H311', description: 'Tóxico en contacto con la piel'}, 
			{value: 'H331', description: 'Tóxico en caso de inhalación'}, 
			{value: 'H314', description: 'Provoca quemaduras graves en la piel y lesiones oculares graves'}, 
			{value: 'H317', description: 'Puede provocar una reacción alérgica en la piel'}
		],
		notas: [
			{value: 'B', description: null},
			{value: 'D', description: 'Ciertas sustancias que pueden experimentar una polimerización o descomposición espontáneas, se comercializan en una forma estabilizada, y así figuran en la parte 3. No obstante, en algunas ocasiones, dichas sustancias se comercializan en una forma no estabilizada. En este caso, el proveedor deberá especificar en la etiqueta el nombre de la sustancia seguido de la palabra «no estabilizada».'}
		],
		concentracionEtiquetado: [
			{concentracion: 'C >= 25 %', etiquetado: 'Corr. cut., 1B; H314'},
			{concentracion: 'C >= 5 %', etiquetado: 'STOT única, 3; H335'},
			{concentracion: 'C >= 0,2 %', etiquetado: 'Sens. cut., 1; H317'},
		],
		descriptions: {
			synonyms: 'Se han incluido otros nombres encontrados en la normativa o en las bases de datos utilizadas para construir la RISCTOX',
			cas: 'Número asignado por el Chemical Abstract Service. Es el sistema de identificación más utilizado a nivel internacional',
			einecs: 'Catálogo europeo de sustancias químicas comercializadas. Dicho inventario establece la lista definitiva de todas las sustancias que en principio se encontraban en el mercado comunitario al 18 de septiembre de 1981.',
			indice: 'El número índice es el número de identificación asignado a la sustancia en el anexo VI del Reglamento 1272/2008 (conocido como CLP) de clasificación, etiquetado y envasado de sustancias y mezclas peligrosas',
			listaNegra: 'La lista negra incluye aquellas sustancias cuyos posibles daños a la salud y al medio ambiente son tan importantes que debemos evitar su uso o presencia en los lugares de trabajo y su vertido al medio ambiente. Estas sustancias, cuya eliminación será prioritaria, son: Cancerígenas, Mutágenas, Tóxicas para la Reproducción, Disruptores Endocrinos, Sensibilizantes, Neurotóxicos, Tóxicas, Persistentes y Bioacumulativas, muy persistent y muy bioacumulativas, Que pueden provocar a largo plazo efectos negativos en el medio ambiente acuático',
			clp: 'Incluye las indicaciones de peligro (frases H) asignadas a las sustancias incluidas en el Anexo VI del Reglamento 1272/2008 (CLP). Los grupos de sustancias incluidos en el Reglamento 1272/2008, Ej. compuestos de berilio, se han desglosado en la relación de sustancias que forman este grupo y se les han asignado las frases H y los pictogramas y palabras de advertencia pertenecientes al grupo.	Las sustancias que pertenecen a 2 grupos, Ej. cromato de mercurio, se han clasificado asignando las frases de mayor riesgo de cada grupo o sus combinaciones, según los criterios del Reglamento CLP.',

		}
	}
})