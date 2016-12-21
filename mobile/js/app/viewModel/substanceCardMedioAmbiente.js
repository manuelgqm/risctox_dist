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
			this.isToxicaAgua = ko.computed( () => inLists(['directiva_aguas', 'alemana'], featuredLists) );
			this.isDirectivaAguas = ko.computed( () =>
				inLists(['directiva_aguas'], featuredLists)
				|| this.directiva_aguas()
			, this);
			this.isPrioritaria = ko.computed( () => inLists(['sustancias_prioritarias'], featuredLists) );
			this.isMedioAmbienteAlemania = ko.computed( () =>
				this.clasif_mma()
				&& !isNaN(!isNaN(this.clasif_mma()[0].key()))
			, this);
		};

		return { viewModel: viewModel, template: template };
	}
);