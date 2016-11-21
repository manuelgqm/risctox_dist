define([
	'app/viewModel/viewModel',
	'text!app/view/substanceCard.html',
	'app/model/Substance',
	'app/viewModel/page',
	'css!app/view/style/substanceCard',
	'css!app/view/style/layout.css'
], function(ViewModel, view, SubstanceModel, pageViewModel){
	module = {
		run: function(){
			var substanceCard = 
				{ domId: 'card'
				, section: module.section
				, hasVlaEstado: function() { return hasItems(this.valoresLimiteAmbiental, "estado") }

				, hasVlaEd: function()
					{
					var ed_mg_m3Valued = this.valoresLimiteAmbiental
						.map(element => element.ed_mg_m3)
						.filter(element => element.length != 0)
						.length;
					var ed_ppmValued = this.valoresLimiteAmbiental
						.map(element => element.ed_ppm)
						.filter(element => element.length != 0)
						.length;
					return ed_mg_m3Valued + ed_ppmValued > 0;
					}

				, hasVlaEc: function()
					{
					var ec_mg_m3Valued = this.valoresLimiteAmbiental
						.map(element => element.ec_mg_m3)
						.filter(element => element.length != 0)
						.length;
					var ec_ppmValued = this.valoresLimiteAmbiental
						.map(element => element.ec_ppm)
						.filter(element => element.length != 0)
						.length;
					return ec_mg_m3Valued + ec_ppmValued > 0;
					}

				, hasVlaNotas: function()
					{ return this.valoresLimiteAmbiental
						.map(element => element.notas)
						.length != 0;
					}

				, hasEinecs: function()
					{ return (
						this.num_ce_einecs 
						&& !beginsWithChar(this.num_ce_einecs, 4) 
						&& !this.num_ce_elincs 					
						)
					}

				, hasElincs: function()
					{ return ( 
						this.num_ce_elincs 
						&& (beginsWithChar(this.num_ce_einecs, 4) 
							|| !this.num_ce_einecs) 
						)
					}

				, hasVlbIb: function() { return hasItems(this.valoresLimiteBiologico, "indicador") }

				, hasVlbValor: function() { return hasItems(this.valoresLimiteBiologico, "valor") }

				, hasVlbMomento: function() { return hasItems(this.valoresLimiteBiologico, "momento") }

				, hasVlbNotas: function()
					{ return this.valoresLimiteBiologico
						.map(element => element.notas).length != 0;	
					}
				, hasClassificationOrLabelingRd1272: function()
					{ return this.pictogramasRd.lengt 
					|| this.clasificacionesRd1272.length
					|| this.notas_rd1272.length
					|| this.concentracionEtiquetadoRd1272.length
					}
				, sectionUrl: function(sectionName)
					{ return '#/card/' + this.id.toString() + '/' + sectionName }
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

			var beginsWithChar = function (string, beginChar) {
				return !string ? false : string.substring(0, 1) == beginChar;
			}

			var hasItems = function(list, elementName) {
				return list
					.map(element => element[elementName])
					.filter(element => element.length != 0)
					.length != 0; 
			}

			return substanceCard;
		}
	}

	return module;
});