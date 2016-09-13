define(['Server'], function(Server){
	return function(id){
		var self = this;
		this.id = id;
		this.load = function(){
			return new Server("substance").request({
				substanceId: self.id
			});
		};
	}
});