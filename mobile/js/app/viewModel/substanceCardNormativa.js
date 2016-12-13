define(
	[ 'knockout'
	, 'text!app/view/substanceCardNormativa.html'
	], function(ko, template){
		function viewModel(card){
			Object.assign(this, card.identificacion);
			var featuredLists = card.featuredLists();
			var inList = function (list1, list2) {
				if (!list1 || !list2)
					return false;
				return list2.indexOf(list1) != -1;
			};
			this.isRestringida = ko.computed( () => inList('restringidas', featuredLists));
			this.isProhibidaEmbarazadas = ko.computed( () => inList('prohibidas_embarazadas', featuredLists));
			this.isProhibidaLactantes = ko.computed( () => inList('prohibidas_lactantes', featuredLists));
		};


		return { viewModel: viewModel, template: template };
	}
);