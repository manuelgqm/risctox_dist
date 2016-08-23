define(['knockout',
		'app/viewModel/substance-card',
		'text!./template/substance-card.html',
		'css!./style/substance-card'
		], 
	function(ko, viewModel, template){
		return {
			render: function(domID){
				document.body.innerHTML = template;
				ko.applyBindings(viewModel, document.getElementById(domID));
			}
		}
	}
);