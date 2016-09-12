require.config({
	baseUrl: 'js/lib',
	paths: {
		'spec': '../spec',
		'app': '../app',
		'jasmine': ['jasmine-2.5.0/jasmine'],
		'jasmine-html': ['jasmine-2.5.0/jasmine-html'],
		'jasmine-boot': ['jasmine-2.5.0/boot'],
		'jasmine-ajax': ['jasmine-2.5.0/mock-ajax'],
		'jquery': 'jquery-3.1.0.min'
	},
	shim: {
		'jasmine-html': {deps : ['jasmine', 'css!jasmine-2.5.0/jasmine.css']},
		'jasmine-ajax': {deps: ['jasmine']},
		'jasmine-boot': {deps : ['jasmine', 'jasmine-html']}
	},
	map: {'*': {'css' : 'require-css/css.min'}}
});
define(['jasmine-boot'], function(){
	require(['spec/substance'], function(){
		window.onload();
	});
});