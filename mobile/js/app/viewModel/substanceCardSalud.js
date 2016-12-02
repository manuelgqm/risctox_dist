define(
	[ 'knockout'
	, 'text!app/view/substanceCardSalud.html'
	], function(ko, template){
		function viewModel(card){
			Object.assign(this, card.salud);
			this.featuredLists = card.featuredLists;
			this.isCancerigenoIarc = this.inList('cancer_iarc');
			this.isDisruptor = this.inList("de");
			this.isNeurotoxico = function(){
				var neurotoxicosLists = 
					[ "neurotoxico"
					, "neurotoxico_rd"
					, "neurotoxico_danesa"
					, "neurotoxico_nivel"
					]
				return inLists(neurotoxicosLists, this.featuredLists()) || this.efecto_neurotoxico == 'OTOTÓXICO';
			}

			var inLists = function(list1, list2){
				var ocurrences = list1.filter( element1 => 
					list2.filter( element2 =>
						element1 == element2
					).length != 0
				).length;
				return ocurrences != 0;	
			}
		};

		viewModel.prototype.isCancerigeno = function(){
			var cancerigenosLists = 
				[ 'cancer_rd'
				, 'cancer_danesa'
				, 'cancer_iarc'
				, 'cancer_otras'
				, 'cancer_mama' ]
			var inCancerigenosLists = cancerigenosLists.filter( 
				cancerigeno => 
					this.featuredLists().filter( featured =>
						cancerigeno == featured
					).length != 0
				).length != 0;
			return inCancerigenosLists;
		};

		viewModel.prototype.inList = function (listName) {
			return this.featuredLists().indexOf(listName) != -1;
		};

		return { viewModel: viewModel, template: template };
	}
);