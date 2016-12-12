define(
	[ 'knockout'
	, 'knockout-mapping'
	, 'app/viewModel/viewModel'
	, 'text!app/view/substanceCard.html'
	, 'app/model/Substance'
	, 'app/model/SubstanceSalud'
	, 'Server'
	, 'css!app/view/style/substanceCard'
	, 'css!app/view/style/layout.css'
], function(ko, mapping, ViewModel, cardView, SubstanceModel, SubstanceSaludModel, Server){
	return function(args){
		Object.assign(ko, mapping);
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
			}

		Object.assign(card, new ViewModel(card, cardView));

		ko.computed(function() {
			section = card[card.section()];
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
				};
			};
			card[card.section()] = load(card.section(), card.substanceId);
		}, this);

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