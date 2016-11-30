define(
	[ 'knockout'
	, 'text!app/view/substanceCardNormativa.html'
	], function(ko, template){
		function viewModel(card){
			Object.assign(this, card.identificacion);
			this.isRestringida = card.featuredLists().indexOf('restringidas') != -1;
			this.isProhibidaEmbarazadas = card.featuredLists().indexOf('prohibidas_embarazadas') != -1;
			this.isProhibidaLactantes = card.featuredLists().indexOf('prohibidas_lactantes') != -1;
		};

		return { viewModel: viewModel, template: template };
	}
);