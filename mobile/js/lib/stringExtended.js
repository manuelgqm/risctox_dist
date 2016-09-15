define(function(){
	String.prototype.decodeHtmlEntity = function(){
		return this.replace(/&#(\d+);/g, function(match, dec) {
		    return String.fromCharCode(dec);
		});
	};
});