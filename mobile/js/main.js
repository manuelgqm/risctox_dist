requirejs.config({
	baseUrl: 'js/lib',
	shim: {'bootstrap': { deps: ['jquery', 'css!bootstrap-3.3.7/css/bootstrap.min.css'] }},
	paths: { 
		app: '../app', 
		knockout: 'knockout-3.4.0', 
		bootstrap: 'bootstrap-3.3.7/js/bootstrap', 
		jquery: 'jquery-3.1.0.min',
		lodash: 'lodash.min'
	}	,
	map: {'*': {'css' : 'require-css/css.min'}}
});
define(['app/viewModel/substanceFinder', 'css!app/view/style/layout'], function( substanceFinderViewModel ){
	substanceFinderViewModel.render('finder');
});