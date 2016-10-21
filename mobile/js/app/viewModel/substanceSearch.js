define(['app/viewModel/ViewModel'
		, 'text!app/view/substanceSearch.html'
		, 'Server'
], function(ViewModel, view, Server){
	'use strict';
	var module = {
		run: function(){
			var search = {
				domId : "search"
				, name: this.name
				, code: this.code
				, results: []
			};
			Object.assign(search, new ViewModel(search, view));

			var returnResults = function(records, search){
				search.results = records;
				search.render();
				search.bind();
			};

			var requestServer = (function(search){
				var ajaxRequest = new Server("substance").request({
					name: search.name
					, code: search.code
					, action: "search"
				}).done( output => returnResults(output.data.records, search) );

				return ajaxRequest;
			})(search);
			
			search.select = current =>	window.location = "#/card/" + current.id;

			return search;
		},


	}
	return module;
});