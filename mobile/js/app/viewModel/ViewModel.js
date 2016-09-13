define([
	'knockout', 
	'bootstrap'
], function(){
	return function(viewModel, template){
		this.render = function(domId){
			document.body.innerHTML = template;
		},
		this.bind = function(){
			domNode = document.getElementById(viewModel.domId);
			ko.applyBindings(viewModel, domNode);
		}
	}
});