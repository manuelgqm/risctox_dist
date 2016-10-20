define(['app/viewModel/ViewModel'
		,'text!app/view/substanceSearch.html'
		, 'Server'
], function(ViewModel, view, Server){
	'use strict';
	var module = {
		run: function(){
			var search = {
				domId : "search",
				name: this.name,
				code: this.code
			};
			Object.assign(search, new ViewModel(search, view));

			var find = function(name, code){
				ajaxRequest = new Server("substance").request({
					name: name
					, code: code
					, action: "search"
				});
				return ajaxRequest;
			};

			find(this.name, this.code);
			search.render();
			search.bind();
			return search;
		},

	}
	return module;
});