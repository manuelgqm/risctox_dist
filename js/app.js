requirejs.config({
	urlArgs: 'v=1.0',
	baseUrl: 'js/lib',
	shim : {
		"bootstrap" : { "deps" :['jquery-1.12.0.min', 'css!bootstrap-3.3.6-dist/css/bootstrap.min'] }
	},
	paths: { 
		app: '..',
		bootstrap: 'bootstrap-3.3.6-dist/js/bootstrap.min',
		'underscore-string': 'underscore.string.min',
		knockout: 'knockout-3.4.0'
	},
	map: {
		'*': {
			'css': 'require-css.min'
		}
	}
});
define( [ 'app/viewModel/model1' ], function( Model ){
	var model = new Model( $( '#main' ) );
} );