define([
	'knockout',
	'app/viewModel/ViewModel',
	'text!app/view/substanceFinder.html', 
], function(ko, ViewModel, view){
	var module = {
		run: function(){
			var Finder = function(domId){
				this.domId = domId;
				this.name = ko.observable('');
				this.code = ko.observable('');
				this.runFinder = function () {
					var self = this;
					var doSearch = () => window.location = "#/search/" + this.name() + "/" + this.code();
					var getValidationMessages = function(){
						var result = [];
						if (0 < self.name().length && self.name().length < 3)	
							result.push("El nombre de la sustancia debe tener al menos 3 caracteres");
						if (self.code().length == 0 && self.name().length < 3)
							result.push("Al menos uno de los campos ha de estar completado");

						return result;
					}
					var notify = messages => messages.map( message => alert(message));
					var validationMessages = getValidationMessages();

					return !validationMessages.length ? doSearch() : notify(validationMessages);
				}

				return this;
			}

			var finder = new Finder('finder');
			Object.assign(finder, new ViewModel(finder, view));
			finder.render();
			finder.bind();
			return finder;
		}
	}

	return module;
});