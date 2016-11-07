define(
	[ 'app/viewModel/ViewModel'
	, 'text!app/view/substanceSearch.html'
	, 'Server'
	, 'css!app/view/style/layout'
	]
, function(ViewModel, view, Server){
	'use strict';
	var module = {
		run: function(){
			var search = {
				domId : "search"
				, name: this.name
				, code: this.code
				, results: []
				, overflow: false
			};

			Object.assign
				( search
				, new ViewModel(search, view)
				, { select : current => showCard(current.id) }
				);

			new Server("substance").request(
				{ name: search.name
				, code: search.code
				, action: "search"
				}
			).done( 
				output => (output.data.records.length == 1) 
					? showCard(output.data.records[0].id)
					: showResults(output.data.records, search)
			);

			var showResults = function(records, search){
				const resultsLimit = 100;
				search.overflow = (records.length > resultsLimit)
					? true
					: false
				search.results = records;
				search.render();
				search.bind();
			};

			var showCard = substanceId => window.location = "#/card/" + substanceId;

			return search;
		},

	};
	return module;
});