class Hydrogen_cyanide
  attr_reader :name, :synonyms, :cas_num, :ce_einecs_num, :groups, :uses
  attr_reader :icsc_nums, :rd_num, :molecular_formula

  def initialize(browser)
    @browser = browser
    @name = "hydrogen cyanide"
    @synonyms = ["hydrocyanic acid"]
    @cas_num = "74-90-8"
    @ce_einecs_num = "200-821-6"
    @groups = ["cyanides", "cyanides"]
    @uses = ["pesticide"]
    @icsc_nums = ["0492"]
    @icsc_nums_links = ["http://www.ilo.org/dyn/icsc/showcard.display?p_lang=en&p_card_id=0492"]
    @rd_num = "006-006-00-X"
    @molecular_formula = "CHN"
  end

  def go
    @browser.goto("http://localhost:8081/en/dn_risctox_ficha_sustancia.asp?id_sustancia=953980")
  end
end
