define(['Server'], function(Server){
	return function(id){
		var self = this;
		this.id = id;
		this.load = function(){
			new Server("substance").request({
				substanceId: self.id
			}).done(function(response){
				console.log(response)
			})
		};
		this.load();
	}
});