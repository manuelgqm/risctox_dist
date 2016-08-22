requirejs.config({
	baseUrl: 'js/lib',
	paths: { app: '../app', knockout: 'knockout-3.4.0'}	
});
define(['app/view/substance-finder'], function( view ){
	view.render();
});