define(
	[ 'knockout'
	, 'text!app/view/substanceCardNormativa.html'
	], function(ko, template){
		function viewModel(card){
			Object.assign(this, card.identification);
		};

		return { viewModel: viewModel, template: template };
	}
);