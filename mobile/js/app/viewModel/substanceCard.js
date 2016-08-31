define([
	'lodash',
	'app/viewModel/viewModel',
	'text!app/view/substanceCard.html',
	'app/model/substanceCard',
	'app/viewModel/page',
	'css!app/view/style/substanceCard'
], function(_, ViewModel, substanceCardView, substanceCardModel, pageViewModel){
	var substanceCardViewModel = {
		domId: 'card',
		showPage: function(){
			pageViewModel.render('page');
		}
	};
	_.assign(substanceCardViewModel, 
		substanceCardModel, 
		new ViewModel(substanceCardViewModel, substanceCardView)
	);

	return substanceCardViewModel;
})