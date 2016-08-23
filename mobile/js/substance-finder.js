requirejs.config({
	baseUrl: 'js/lib',
	paths: { app: '../app', knockout: 'knockout-3.4.0'}	,
	map: {'*': {'css' : 'require-css/css.min'}}
});
define(['app/view/substance-finder'], function( view ){
	view.render('finder');
});