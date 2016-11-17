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
				, hasEinecs: function()
					{ return this.num_ce_einecs && this.num_ce_einecs.substring(0, 1) != 4 && !this.num_ce_elincs }
				, hasElincs: function()
					{ return this.num_ce_elincs && (this.num_ce_einecs.substring(0, 1)  == 4 || !this.num_ce_einecs) }
				, hasVlbIb: function()
					{ return this.valoresLimiteBiologico.map(element => element.indicador).filter(element => element.length != 0).length != 0; }
				, hasVlbValor: function()
					{ return this.valoresLimiteBiologico.map(element => element.valor).filter(element => element.length != 0).length != 0; }
				, hasVlbMomento: function()
					{ return this.valoresLimiteBiologico.map(element => element.momento).filter(element => element.length != 0).length != 0; }
				, hasVlbNotas: function()
					{ return this.valoresLimiteBiologico.map(element => element.notas).length != 0;	}
				, hasClassificationOrLabelingRd1272: function()
					{ return this.pictogramasRd.lengt 
					|| this.clasificacionesRd1272.length
					|| this.notas_rd1272.length
					|| this.concentracionEtiquetadoRd1272.length
					}
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