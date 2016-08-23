requirejs.config({
	baseUrl: 'js/lib',
	paths: { app: '../app', knockout: 'knockout-3.4.0'}	,
	map: {'*': {'css' : 'require-css/css.min'}}
});
define(['app/view/substance-finder', 'css!app/view/style/layout'], function( view ){
	view.render('finder');
});