define([
	'lodash',
	'app/viewModel/viewModel',
	'text!app/view/template/substanceFinder.html', 
	'app/viewModel/substanceCard',
], function(_, View, substanceFinderTemplate, cardView){
	var substanceFinderViewModel = {
		id: 957597,
		findSubstance: function(){
			cardView.render('card');
		}
	};
	_.assign(substanceFinderViewModel, new View(substanceFinderViewModel, substanceFinderTemplate));
	return substanceFinderViewModel;
});