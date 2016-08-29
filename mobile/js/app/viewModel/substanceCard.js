define([
	'lodash',
	'app/viewModel/viewModel',
	'app/view/page',
	'text!app/view/substanceCard.html',
	'app/model/substanceCard',
	'css!app/view/style/substanceCard'
], function(_, ViewModel, pageView, substanceCardView, substanceCardModel){
	var substanceCardViewModel = {
		showPage: function(){
			pageView.render('card');
		}
	};
	_.assign(substanceCardViewModel, 
		substanceCardModel, 
		new ViewModel(substanceCardViewModel, substanceCardView)
	);

	return substanceCardViewModel;
})