define(
	[ 'knockout'
	, 'Server'
	, 'text!app/view/substanceCardSalud.html' ]
	, function(ko, Server, template){
		function viewModel(card){
			Object.assign(this, card.salud);
			var self = this;
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
			this.isCancerOtras = ko.computed( () => inLists(['cancer_otras'], featuredLists));
			this.isEnfermedad = ko.computed( () => inLists(['eepp'], featuredLists));
			this.loadCancerOtras = function(){
				if (this.categorias_cancer_otras()) 
					return false;
				new Server("substance").request(
					{ substanceId: card.substanceId
					, action: "findCancerOtras"	}
				).done(function(output){
					ko.fromJS(output.data, self);
				})
			};
			this.loadEnfermedades = function(){
				if (this.enfermedades_profesionales())
					return false;
				new Server("substance").request(
					{ substanceId: card.substanceId
					, action: "findEnfermedades" }
				).done(function(output){
					ko.fromJS(output.data, self);
				})
			};
			this.getCollapseEnfermedadId = function(args){ 
				return 'enfermedadesId' + args.id();
			};
			
		};

		return { viewModel: viewModel, template: template };
	}
);