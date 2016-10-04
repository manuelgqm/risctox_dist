define([
	'app/viewModel/viewModel',
	'text!app/view/substanceCard.html',
	'app/model/Substance',
	'app/viewModel/page',
	'css!app/view/style/substanceCard'
], function(ViewModel, substanceCardView, SubstanceModel, pageViewModel){
	module = {
		run: function(){
			var substanceCardViewModel = {domId: 'card'};
			var substanceId = 957597;
			var substance = new SubstanceModel(substanceId);
			Object.assign(substanceCardViewModel, substance, new ViewModel(substanceCardViewModel, substanceCardView));
			substance.load().done(function(output){
				Object.assign(substanceCardViewModel, 
					output.data, 
					new ViewModel(substanceCardViewModel, substanceCardView)
				);
				substanceCardViewModel.render();
				substanceCardViewModel.bind();
			});	

			return substanceCardViewModel;
		}
	}

	return module;
});