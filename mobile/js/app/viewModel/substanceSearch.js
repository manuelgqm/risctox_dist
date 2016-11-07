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
			const pipe = (...fns) => (x) => fns.reduce((prev, func) => func(prev), x);
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
					: pipe
						( setResults(output.data.records)
						, showResults()
						) 
			);
			
			var showCard = substanceId => window.location = "#/card/" + substanceId;
			var setResults = function(records){
				const resultsLimit = 100;
				search.results = records;
				search.overflow = (records.length > resultsLimit)
					? true
					: false
			};
			var showResults = function(){
				console.log("de")
				search.render();
				search.bind();
			};

			return search;
		},

	};
	return module;
});