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
		};

		return { viewModel: viewModel, template: template };
	}
);