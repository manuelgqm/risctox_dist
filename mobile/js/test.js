require.config({
	baseUrl: 'js/lib',
	paths: {
		'jasmine': ['jasmine-2.5.0/jasmine'],
		'jasmine-html': ['jasmine-2.5.0/jasmine-html'],
		'jasmine-boot': ['jasmine-2.5.0/boot'],
		'spec': '../spec',
		'app': '../app'
	},
	shim: {
		'jasmine-html': {deps : ['jasmine', 'css!jasmine-2.5.0/jasmine.css']},
		'jasmine-boot': {deps : ['jasmine', 'jasmine-html']}
	},
	map: {'*': {'css' : 'require-css/css.min'}}
});
define(['jasmine-boot'], function(){
	require(['spec/Substance'], function(){
		window.onload();
	});
});