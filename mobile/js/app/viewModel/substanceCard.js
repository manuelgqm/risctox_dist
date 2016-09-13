define([
	'lodash',
	'app/viewModel/viewModel',
	'text!app/view/substanceCard.html',
	'app/model/substanceCard',
	'app/model/Substance',
	'app/viewModel/page',
	'css!app/view/style/substanceCard'
], function(_, ViewModel, substanceCardView, substanceCardModel, SubstanceModel, pageViewModel){
	var substanceCardViewModel = {domId: 'card'};
	var substanceId = 957597;
	var substance = new SubstanceModel(substanceId);
	_.assign(substanceCardViewModel, 
		substanceCardModel, 
		new ViewModel(substanceCardViewModel, substanceCardView)
	);

	return substanceCardViewModel;
})