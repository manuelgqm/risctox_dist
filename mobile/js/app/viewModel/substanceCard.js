define(
	[ 'knockout'
	, 'knockout-mapping'
	, 'Server'
	, 'app/viewModel/viewModel'
	, 'text!app/view/substanceCard.html'
	, 'app/model/Substance'
	, 'app/model/SubstanceSalud'
	, 'app/model/SubstanceMedioAmbiente'
	, 'css!app/view/style/substanceCard'
	, 'css!app/view/style/layout.css'
], function(ko, mapping, Server, ViewModel, template, SubstanceModel, SubstanceSaludModel, SubstanceMedioAmbienteModel) {
	'use strict';
	return function(args){
		Object.assign(ko, mapping);
		var inLists = function(list1, list2) {
			if (!list1 || !list2)
				return false;
			return list1.filter( element => 
				list2.indexOf(element) != -1
			).length != 0;
		};

		var card = 
			{ domId : 'card'
			, section : ko.observable(args.section || 'identificacion')
			, isSection : function(currentSection){ return this.section() == currentSection }
			, substanceId : args.id
			, identificacion : {}
			, salud: {}
			, normativa: {}
			, medioAmbiente: {}
			, setSection: function(section) 
				{ this.section(section) }
			};
		Object.assign(card, new ViewModel(card, template));
		ko.fromJS(card);

		ko.computed(function() {
			var section = card[card.section()];
			if (Object.keys(section).length) {
				return true;
			};

			var load = function(sectionName, substanceId){
				var result = {};
				Object.assign(result
					, ko.fromJS(
						createSection(sectionName)
						)
					);

				new Server("substance").request(
					{ substanceId: substanceId
					, action: "findSection"
					, section: sectionName
					}).done(function(output){
						ko.fromJS(output.data, result);
					});

				return result;
			};

			var createSection = function(sectionName){
				switch(sectionName){
					case("salud") : return new SubstanceSaludModel();
					case("identificacion") : return new SubstanceModel();
					case("medioAmbiente") : return new SubstanceMedioAmbienteModel();
				};
			};
			card[card.section()] = load(card.section(), card.substanceId);
		}, this);

		card.hasSalud = ko.computed( () =>
			inLists(
				[ "cancer_iarc"
				, "de"
				, 'cancer_rd'
				, 'cancer_danesa'
				, 'cancer_iarc'
				, 'cancer_otras'
				, 'cancer_mama'
				, "neurotoxico"
				, "neurotoxico_rd"
				, "neurotoxico_danesa"
				, "neurotoxico_nivel" ]
			, card.identificacion.featuredLists()) 
			|| this.efecto_neurotoxico == 'OTOTÓXICO'
		);

		card.hasNormativa = ko.computed( () => 
			inLists(
				[ "restringias"
				, "prohibidas_embarazadas"
				, "prohibidas_lactentes" ]
			, card.identificacion.featuredLists())
		);

		card.hasMedioAmbiente = ko.computed( () =>
			inLists(['tpb'], card.identificacion.featuredLists())
		);

		var registerComponent = function(componentName, viewModelName){
			if (ko.components.isRegistered(componentName)) {
				return false;
			};
			ko.components.register(componentName, 
				{ require: 'app/viewModel/' + viewModelName }
			);
			return true;
		};

		registerComponent('identificacion', 'substanceCardIdentificacion');
		registerComponent('salud', 'substanceCardSalud');
		registerComponent('normativa', 'substanceCardNormativa');
		registerComponent('medioAmbiente', 'substanceCardMedioAmbiente');

		card.render();
		card.bind();
	};
});