define(
	[ 'knockout'
	, 'text!app/view/substanceCardNormativa.html'
	], function(ko, template){
		function viewModel(card){
			Object.assign(this, card.identificacion);
			this.isRestringida = card.featuredLists().indexOf('restringidas') != -1;
		};

		return { viewModel: viewModel, template: template };
	}
);