define([
	'lodash',
	'app/viewModel/ViewModel',
	'text!app/view/substanceFinder.html', 
	'app/model/substanceFinder',
	'app/viewModel/substanceCard',
], function(_, ViewModel, substanceFinderView, substanceFinderModel, substanceCardViewModel){
	var substanceFinderViewModel = {
		findSubstance: function(){
			substanceCardViewModel.render('card');
		}
	};
	_.assign(substanceFinderViewModel, 
		substanceFinderModel,
		new ViewModel(substanceFinderViewModel, substanceFinderView)
	);
	return substanceFinderViewModel;
});