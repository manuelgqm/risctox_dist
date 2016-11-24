define([
	'knockout',
	'app/viewModel/viewModel',
	'text!app/view/substanceCard.html',
	'css!app/view/style/substanceCard',
	'css!app/view/style/layout.css'
], function(ko, ViewModel, view){
	module = {
		run: function(){
			var substanceCard = 
				{ domId: 'card'
				, section: module.section || 'identificacion'
				, substanceId: module.id
				, sectionUrl: function(sectionName)
					{ return '#/card/' + this.substanceId.toString() + '/' + sectionName }
				};
			Object.assign(substanceCard, new ViewModel(substanceCard, view));
			if (!ko.components.isRegistered('identificacion')) {
				ko.components.register('identificacion', { require: 'app/viewModel/substanceCardIdentificacion' });
			}
			substanceCard.render();
			substanceCard.bind();

			return substanceCard;
		}
	}

	return module;
});