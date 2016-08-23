define(['app/viewModel/substance-finder', 
		'text!./template/substance-finder.html', 
		'knockout', 'bootstrap'], 
	function(viewModel, template, ko){
		return {
			render: function( domID ){
				document.body.innerHTML = template;
				ko.applyBindings(viewModel, document.getElementById(domID));
			}
		}
	}
);