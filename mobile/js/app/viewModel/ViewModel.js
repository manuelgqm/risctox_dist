define([
	'knockout', 
	'bootstrap'
], function(ko){
	return function(viewModel, template){
		this.viewModel = viewModel;
		this.template = template;
		this.render = function(domId){
			document.body.innerHTML = this.template;
			ko.applyBindings(this.viewModel, document.getElementById(domId));
		}
	}
});