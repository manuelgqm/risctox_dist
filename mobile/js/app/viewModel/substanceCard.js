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

		var loadSection = ko.computed(function() {
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

		ko.components.register('identificacion', { require: 'app/viewModel/substanceCardIdentificacion' });
		ko.components.register('salud', { require: 'app/viewModel/substanceCardSalud' });
		ko.components.register('normativa', { require: 'app/viewModel/substanceCardNormativa' });
		ko.components.register('medioAmbiente', { require: 'app/viewModel/substanceCardMedioAmbiente' });

		card.render();
		card.bind();
	};
});