define(['app/model/Substance'], function(Substance){
	describe("A substance", function(){
		it("Must have a id", function(){
			var substance = new Substance(957597);
			expect(substance.id).toEqual(957597);
		});

	});
});