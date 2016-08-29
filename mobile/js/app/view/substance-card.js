define([
	'app/view/View',
	'app/viewModel/substance-card',
	'text!./template/substance-card.html',
	'css!./style/substance-card',
], function(View, viewModel, template){
	var substanceCardView = new View(viewModel, template);
	return substanceCardView;
});