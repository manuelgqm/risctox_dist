define([
	'app/viewModel/ViewModel',
	'text!app/view/substanceSearch.html', 
], function(ViewModel, view){
	var module = {
		run: function(){
			var search = {
				domId : "search",
				name: this.name,
				code: this.code
			};
			Object.assign(search, new ViewModel(search, view));
			search.render();
			search.bind();
			return search;
		}
	}
	return module;
});