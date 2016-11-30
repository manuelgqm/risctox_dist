define(
	[ 'knockout'
	, 'text!app/view/substanceCardSalud.html'
	], function(ko, template){
		function viewModel(card){
			this.featuredLists = card.featuredLists;
			Object.assign(this, card.salud);
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

		return { viewModel: viewModel, template: template };
	}
);