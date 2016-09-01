define([
	'lodash',
	'app/viewModel/ViewModel',
	'text!app/view/substanceFinder.html', 
	'app/model/substanceFinder'
], function(_, ViewModel, substanceFinderView, substanceFinderModel){
	var substanceFinderViewModel = {domId: 'finder'};
	_.assign(substanceFinderViewModel, 
		substanceFinderModel,
		new ViewModel(substanceFinderViewModel, substanceFinderView)
	);
	return substanceFinderViewModel;
});