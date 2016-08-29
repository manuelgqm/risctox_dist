define([
	'lodash',
	'app/viewModel/viewModel',
	'app/view/page',
	'text!app/view/template/substanceCard.html',
	'app/model/substanceCard',
	'css!app/view/style/substanceCard'
], function(_, ViewModel, pageView, substanceCardTemplate, substanceCardModel){
	var substanceCardViewModel = {
		showPage: function(){
			pageView.render('card');
		}
	};
	_.assign(substanceCardViewModel, 
		substanceCardModel, 
		new ViewModel(substanceCardViewModel, substanceCardTemplate)
	);

	return substanceCardViewModel;
})