define([
	'knockout', 
	'bootstrap'
], function(ko){
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