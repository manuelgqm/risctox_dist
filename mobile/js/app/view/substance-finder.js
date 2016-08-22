define(['app/viewModel/substance-finder', 
		'text!./template/substance-finder.html', 
		'knockout'], 
	function(viewModel, template, ko){
		return {
			render: function( domElement ){
				document.body.innerHTML = template;
				ko.applyBindings(viewModel);
			}
		}
	}
);