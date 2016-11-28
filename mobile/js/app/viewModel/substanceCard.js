define(
	[ 'knockout'
	, 'knockout-mapping'
	, 'app/viewModel/viewModel'
	, 'text!app/view/substanceCard.html'
	, 'app/model/substance'
	, 'Server'
	, 'css!app/view/style/substanceCard'
	, 'css!app/view/style/layout.css'
], function(ko, mapping, ViewModel, cardView, SubstanceModel, Server){
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
		Object.assign(card.identificacion, loadIdentificacion(card.substanceId));

		function loadIdentificacion(substanceId){
			var result = {};
			Object.assign(result, new SubstanceModel(substanceId));
			Object.assign(result, ko.fromJS(result));

			result.load().done(function(output){
				ko.fromJS(output.data, result);
			});

			return result;
		}

		loadSection = ko.computed(
			function(){
				section = card[card.section()];
				if (Object.keys(section).length) {
					return true;
				}
				Object.assign(section, 
					{ grupo_iarc: null
					, volumen_iarc: null
					, notas_iarc: null
					});
				Object.assign(section, ko.fromJS(section));
				new Server("substance").request(
					{ substanceId: card.substanceId
					, action: "findSalud"
					}).done(function(output){
						ko.fromJS(output.data, section);
					});
			}
		, this);

		ko.components.register('identificacion', { require: 'app/viewModel/substanceCardIdentificacion' });
		ko.components.register('salud', { require: 'app/viewModel/substanceCardSalud' });
		ko.components.register('normativa', { require: 'app/viewModel/substanceCardNormativa' });
		ko.components.register('medioAmbiente', { require: 'app/viewModel/substanceCardMedioAmbiente' });

		card.render();
		card.bind();
	};
});