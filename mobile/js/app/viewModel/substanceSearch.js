define(['app/viewModel/ViewModel'
		,'text!app/view/substanceSearch.html'
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

			var requestServer = function(search){
				var ajaxRequest = new Server("substance").request({
					name: search.name
					, code: search.code
					, action: "search"
				});
				return ajaxRequest;
			};

			var show = function(search){
				search.render();
				search.bind();
			};

			var setResults = function(search, results){
				search.results = results;
				return search;
			}

			requestServer(search).done( function(output){
				setResults(search, output.data.records);
				show(search);
			});

			return search;
		},

	}
	return module;
});