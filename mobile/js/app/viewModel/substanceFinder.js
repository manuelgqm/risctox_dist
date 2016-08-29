define([
	'lodash',
	'app/viewModel/ViewModel',
	'text!app/view/template/substanceFinder.html', 
	'app/model/substanceFinder',
	'app/viewModel/substanceCard',
], function(_, ViewModel, substanceFinderTemplate, substanceFinderModel, substanceCardViewModel){
	var substanceFinderViewModel = {
		findSubstance: function(){
			substanceCardViewModel.render('card');
		}
	};
	_.assign(substanceFinderViewModel, 
		substanceFinderModel,
		new ViewModel(substanceFinderViewModel, substanceFinderTemplate)
	);
	return substanceFinderViewModel;
});