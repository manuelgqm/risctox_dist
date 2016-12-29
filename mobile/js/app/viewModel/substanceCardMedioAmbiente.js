define(
	[ 'knockout'
	, 'text!app/view/substanceCardMedioAmbiente.html'
	], function(ko, template){
		function viewModel(card){
			Object.assign(this, card.medioAmbiente);
			var featuredLists = card.featuredLists();
			var inLists = function(list1, list2) {
				if (!list1 || !list2)
					return false;
				return list1.filter( element => 
					list2.indexOf(element) != -1
				).length != 0;
			};
			this.isTpb = ko.computed( () =>	inLists(['tpb'], featuredLists) );
			this.isToxicaAgua = ko.computed( () => inLists(['directiva_aguas', 'alemana', 'sustancias_prioritarias'], featuredLists) );
			this.isDirectivaAguas = ko.computed( () => inLists(['directiva_aguas'], featuredLists) );
			this.isPrioritaria = ko.computed( () => inLists(['sustancias_prioritarias'], featuredLists) );
			this.isMedioAmbienteAlemania = ko.computed( () => inLists(['alemana'], featuredLists) );
			this.isContaminanteAire = ko.computed( () =>
				inLists(['ozono', 'clima', 'aire'], featuredLists)
			);
			this.isCalidadAire = ko.computed( () => inLists(['aire'], featuredLists) );
			this.isDanoOzono = ko.computed( () => inLists(['ozono'], featuredLists) );
			this.isCambioClima = ko.computed( () => inLists(['clima'], featuredLists) );
		};

		return { viewModel: viewModel, template: template };
	}
);