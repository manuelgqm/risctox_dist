define([
	'app/view/View',
	'app/viewModel/substance-finder', 
	'text!./template/substance-finder.html', 
], function(View, viewModel, template){
	var substanceFinderView = new View(viewModel, template);
	return substanceFinderView;
});