define(
	[ 'knockout'
	, 'text!app/view/substanceCardNormativa.html'
	], function(ko, template){
		function viewModel(card){
			Object.assign(this, card.identificacion);
			this.featuredLists = card.featuredLists;
			this.isRestringida = this.inList('restringidas');
			this.isProhibidaEmbarazadas = this.inList('prohibidas_embarazadas');
			this.isProhibidaLactantes = this.inList('prohibidas_lactantes');
		};

		viewModel.prototype.inList = function (listName) {
			return this.featuredLists().indexOf(listName) != -1;
		};

		return { viewModel: viewModel, template: template };
	}
);