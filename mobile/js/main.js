define([], function(){
	requirejs.config({
		baseUrl: 'js/lib',
		shim: {'bootstrap': { deps: ['jquery', 'css!bootstrap-3.3.7/css/bootstrap.min.css'] }},
		paths: { 
			app: '../app', 
			router: 'router.min',
			knockout: 'knockout-3.4.0', 
			bootstrap: 'bootstrap-3.3.7/js/bootstrap', 
			jquery: 'jquery-3.1.0.min'
		},
		map: {'*': {'css' : 'require-css/css.min'}}
	});

	require(['router'], function(router){
		router.registerRoutes({
			substanceFinder: {path: '/finder', moduleId: 'app/viewModel/substanceFinder'},
			substanceSearch: {path: '/search/:name/:code', moduleId: 'app/viewModel/substanceSearch'},
			substanceCard: {path: '/card/:id', moduleId: 'app/viewModel/substanceCard'},
			page: {path: '/page', moduleId: 'app/viewModel/page'},
			notFound: {path: '*', moduleId: 'app/viewModel/substanceFinder'}
		});
		router.on('routeload', function(module, args){
			Object.assign(module, args);
			module.run();
		});
		router.init();
	});

});