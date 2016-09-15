define([
	'lodash',
	'app/viewModel/viewModel',
	'text!app/view/substanceCard.html',
	'app/model/Substance',
	'app/viewModel/page',
	'css!app/view/style/substanceCard'
], function(_, ViewModel, substanceCardView, SubstanceModel, pageViewModel){
	var substanceCardViewModel = {domId: 'card'};
	var substanceId = 957597;
	var substance = new SubstanceModel(substanceId);
	_.assign(substanceCardViewModel, substance, new ViewModel(substanceCardViewModel, substanceCardView));
	substance.load().done(function(output){
		_.assign(substanceCardViewModel, 
			output.data, 
			new ViewModel(substanceCardViewModel, substanceCardView)
		);
		substanceCardViewModel.bind();
	});	

	return substanceCardViewModel;
});