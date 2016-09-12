define( ['jquery'], function($){
	function Server(serviceName){
		this.service_url = '../lib/service/' + serviceName + '.asp';
	}
	
	Server.prototype.request = function(params){
		that = this;
		return $.ajax({ 
			url: that.service_url, 
			dataType: "json", 
			async: true, 
			data: params, 
			success: function(output){that.data = output}
		});
	}
	
	return Server;
})