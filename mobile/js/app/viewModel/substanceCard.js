define([
	'app/viewModel/viewModel',
	'text!app/view/substanceCard.html',
	'app/model/Substance',
	'app/viewModel/page',
	'css!app/view/style/substanceCard'
], function(ViewModel, view, SubstanceModel, pageViewModel){
	module = {
		run: function(){
			var substanceCard = 
				{ domId: 'card'
				, hasVlaEstado: function()
					{ return this.valoresLimiteAmbiental.map(element => element.estado).filter(element => element.length != 0).length != 0;	}
				, hasVlaEd: function()
					{
					var ed_mg_m3Valued = this.valoresLimiteAmbiental.map(element => element.ed_mg_m3).filter(element => element.length != 0).length;
					var ed_ppmValued = this.valoresLimiteAmbiental.map(element => element.ed_ppm).filter(element => element.length != 0).length;
					return ed_mg_m3Valued + ed_ppmValued > 0;
					}
				, hasVlaEc: function()
					{
					var ec_mg_m3Valued = this.valoresLimiteAmbiental.map(element => element.ec_mg_m3).filter(element => element.length != 0).length;
					var ec_ppmValued = this.valoresLimiteAmbiental.map(element => element.ec_ppm).filter(element => element.length != 0).length;
					return ec_mg_m3Valued + ec_ppmValued > 0;
					}
				, hasVlaNotas: function()
					{ return this.valoresLimiteAmbiental.map(element => element.notas).length != 0;	}
				};
			var substance = new SubstanceModel(this.id);
			Object.assign
				( substanceCard
				, substance
				, new ViewModel(substanceCard, view)
				);

			substance.load().done(function(output){
				Object.assign
					( substanceCard 
					, output.data
					, new ViewModel(substanceCard, view)
					);
				substanceCard.render();
				substanceCard.bind();
			});

			return substanceCard;
		}
	}

	return module;
});