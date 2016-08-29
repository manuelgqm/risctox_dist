define([
	'app/viewModel/ViewModel', 
	'app/model/page',
	'text!app/view/page.html'
	], 
	function(ViewModel, pageModel, pageView){
		var pageViewModel = {};
		_.assign(pageViewModel, 
			pageModel,
			new ViewModel(pageViewModel, pageView)
		);
		
		return pageViewModel;
	}
);