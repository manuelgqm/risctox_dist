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
				var ajaxRequest = new Server("substance").request({
					name: name
					, code: code
					, action: "search"
				});
				return ajaxRequest;
			};

			var show = function(search){
				search.render();
				search.bind();
			};

			find(this.name, this.code).done( 
				output => show(search)
			);

			return search;
		},

	}
	return module;
});