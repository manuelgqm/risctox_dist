define(['text!./template/substance-finder.html'], function(template){
	return {
		render: function( domElement ){
			document.body.innerHTML = template;
		}
	}
});