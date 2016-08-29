define([
	'lodash',
	'app/view/View',
	'text!app/view/template/substance-finder.html', 
	'app/viewModel/substance-card',
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