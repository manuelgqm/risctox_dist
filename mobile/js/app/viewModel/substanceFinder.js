define([
	'app/viewModel/ViewModel',
	'text!app/view/substanceFinder.html', 
	'app/model/substanceFinder'
], function(ViewModel, substanceFinderView, substanceFinderModel){
	var substanceFinderViewModel = {domId: 'finder'};
	Object.assign(substanceFinderViewModel, 
		substanceFinderModel,
		new ViewModel(substanceFinderViewModel, substanceFinderView)
	);
	return substanceFinderViewModel;
});