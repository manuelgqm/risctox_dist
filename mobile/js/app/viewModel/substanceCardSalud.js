define(
	[ 'knockout'
	, 'text!app/view/substanceCardSalud.html'
	], function(ko, template){
		function viewModel(card){
			Object.assign(this, card.salud);
		};

		return { viewModel: viewModel, template: template };
	}
);