define(
	[ 'knockout'
	, 'text!app/view/substanceCardIdentification.html'
	, 'app/model/substance'
	, 'knockout-mapping'
	], function(ko, template, SubstanceModel, mapping){
		Object.assign(ko, mapping);

		function viewModel(card){
			var self = this;
			Object.assign(this, new SubstanceModel(card.substanceId));
			Object.assign(this, ko.fromJS(this));
			this.load().done(function(output){
				ko.fromJS(output.data, self);
			});

			this.hasEinecs = ko.computed(function()
				{ return (
					this.num_ce_einecs()
					&& !beginsWithChar(this.num_ce_einecs(), 4) 
					&& !this.num_ce_elincs() 					
					)
				}
			, this);

			this.hasElincs = ko.computed(function()
				{ return ( 
					this.num_ce_elincs 
					&& (beginsWithChar(this.num_ce_einecs(), 4) || !this.num_ce_einecs) 
					)
				}
			, this);

			this.hasClassificationOrLabelingRd1272 = ko.computed(function()
				{ return ( 
					this.pictogramasRd()
					|| this.clasificacionesRd1272()
					|| this.notas_rd1272()
					|| this.concentracionEtiquetadoRd1272()
					)
				}
			, this);

			this.hasVlaEstado = ko.computed(function() 
				{ return hasItems(this.valoresLimiteAmbiental(), "estado") }
			, this);

			this.hasVlaEd = ko.computed(function()
				{ return(
					hasItems(this.valoresLimiteAmbiental(), "ed_mg_m3")
					|| hasItems(this.valoresLimiteAmbiental(), "ed_ppm")
					)
				}
			, this);

			this.hasVlaEc = ko.computed(function()
				{ return(
					hasItems(this.valoresLimiteAmbiental(), "ec_mg_m3")
					||hasItems(this.valoresLimiteAmbiental(), "ec_ppm")
					)
				}
			, this);

			this.hasVlaNotas = ko.computed(function()
				{ return hasNotas(this.valoresLimiteAmbiental()) }
			, this);

			this.hasValoresLimiteAmbiental = ko.computed(function()
				{ return (
					this.hasVlaEstado()
					|| this.hasVlaEc()
					|| this.hasVlaEd()
					|| this.hasVlaNotas()
					)
				}
			, this);

			this.hasVlbIb = ko.computed(function() 
				{ return hasItems(this.valoresLimiteBiologico(), "indicador") } 
			, this);

			this.hasVlbValor = ko.computed(function() 
				{ return hasItems(this.valoresLimiteBiologico(), "valor") }
			, this);

			this.hasVlbMomento = ko.computed(function() 
				{ return hasItems(this.valoresLimiteBiologico(), "momento") }
			, this);

			this.hasVlbNotas = ko.computed(function()
				{ return hasNotas(this.valoresLimiteBiologico()) } 
			, this);

			this.hasValoresLimiteBiologico = ko.computed(function()
				{ return (
					this.hasVlbIb()
					|| this.hasVlbValor()
					|| this.hasVlbMomento()
					|| this.hasVlbNotas()
					)
				}
			, this);

			this.hasValoresLimite = ko.computed(function()
				{ return this.hasValoresLimiteAmbiental() || this.hasValoresLimiteBiologico() }
			, this);
			
			return this
		};

		var beginsWithChar = function(string, beginChar) 
			{ return !string ? false : string.substring(0, 1) == beginChar }

		var hasItems = function(list, elementName)
			{ 	
				return list
				.map(element => element[elementName])
				.filter(element => element() ? true : false)
				.length != 0; 
			}

		var hasNotas = function(valoresLimite)
			{ return valoresLimite
				.map(element => hasItems(element.notas(), "key") )
				.filter(element => element ? true : false)
				.length != 0;
			} 

		return { viewModel: viewModel, template: template, syncronous: true };
	}
); 