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
					var doSearch = () => window.location = "#/search/" + this.name() + "/" + this.code();
					var validate = () => this.name().length + this.code().length != 0;
					var notify = () => alert("rellena alguno de los campos");

					return validate() ? doSearch() : notify();
				};

				return this;
			};

			var finder = new Finder('finder');
			Object.assign(finder, new ViewModel(finder, view));
			finder.render();
			finder.bind();
			return finder;
		}
	}

	return module;
});