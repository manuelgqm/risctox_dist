requirejs.config({
	baseUrl: 'js/lib',
	paths: { app: '../app'}	
});
define(['app/view/substance-finder'], function( view ){
	view.render();
});