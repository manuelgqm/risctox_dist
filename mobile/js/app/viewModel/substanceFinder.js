define([
	'knockout',
	'app/viewModel/ViewModel',
	'text!app/view/substanceFinder.html', 
], function(ko, ViewModel, view){
	var module = {
		run: function(){
			var finder = 
				{ domId: "finder"
				, name: ko.observable('')
				, code: ko.observable('')
				, messages: ko.observableArray()
			};

			finder.find = function(){
				if (0 < this.name().length && this.name().length < 3)
					this.messages.push("El nombre de la sustancia debe tener al menos 3 caracteres");
				if (this.code().length == 0 && this.name().length < 3)
					this.messages.push("Al menos uno de los campos ha de estar completado");
				var doSearch = () => window.location = "#/search/" + this.name() + "/" + this.code();

				return this.messages().length ? false : doSearch();
			};

			Object.assign(finder, new ViewModel(finder, view));
			finder.render();
			finder.bind();
			return finder;
		}
	}

	return module;
});