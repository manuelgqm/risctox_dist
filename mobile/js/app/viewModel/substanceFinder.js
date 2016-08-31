define([
	'lodash',
	'app/viewModel/ViewModel',
	'text!app/view/substanceFinder.html', 
	'app/model/substanceFinder',
	'app/viewModel/substanceCard',
], function(_, ViewModel, substanceFinderView, substanceFinderModel, substanceCardViewModel){
	var substanceFinderViewModel = {
		domId : 'finder',
		findSubstance: function(){
			substanceCardViewModel.render();
			substanceCardViewModel.bind();
		}
	};
	_.assign(substanceFinderViewModel, 
		substanceFinderModel,
		new ViewModel(substanceFinderViewModel, substanceFinderView)
	);
	return substanceFinderViewModel;
});