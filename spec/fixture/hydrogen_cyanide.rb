class Hydrogen_cyanide
  attr_reader :name, :synonyms, :cas_num, :ce_einecs_num

  def initialize(browser)
    @browser = browser
    @name = "hydrogen cyanide"
    @synonyms = ["hydrocyanic acid"]
    @cas_num = "74-90-8"
    @ce_einecs_num = "200-821-6"
  end

  def go
    @browser.goto("http://localhost:8081/en/dn_risctox_ficha_sustancia.asp?id_sustancia=953980")
  end
end
