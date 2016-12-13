define(
	[ 'knockout'
	, 'text!app/view/substanceCardSalud.html' ]
	, function(ko, template){
		function viewModel(card){
			Object.assign(this, card.salud);
			var featuredLists = card.featuredLists();
			var inLists = function(list1, list2) {
				if (!list1 || !list2)
					return false;
				return list1.filter( element => 
					list2.indexOf(element) != -1
				).length != 0;
			};
			this.isCancerigenoIarc = ko.computed( () => inLists(['cancer_iarc'], featuredLists));
			this.isDisruptor = ko.computed( () => inLists(['de'], featuredLists));
			this.isNeurotoxico = ko.computed( () => 
				inLists(
					[ "neurotoxico"
					, "neurotoxico_rd"
					, "neurotoxico_danesa"
					, "neurotoxico_nivel" ]
				, featuredLists) 
				|| this.efecto_neurotoxico == 'OTOTÓXICO'
			);
			this.isCancerigeno = ko.computed( () => 
				inLists(
					[ 'cancer_rd'
					, 'cancer_danesa'
					, 'cancer_iarc'
					, 'cancer_otras'
					, 'cancer_mama' ]
				, featuredLists)
			);
			this.isToxicoReproduccion = ko.computed( () => inLists(['tpr'], featuredLists));
			
		};

		return { viewModel: viewModel, template: template };
	}
);