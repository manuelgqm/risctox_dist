class Hydrogen_cyanide
  attr_reader :id
  attr_reader :name, :synonyms, :cas_num, :ce_einecs_num, :groups, :uses
  attr_reader :icsc_nums, :rd_num, :molecular_formula, :concern_trade_union_reasons
  attr_reader :rd1272_symbols, :h_phrases

  def initialize()
    @id = 953980
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
    @concern_trade_union_reasons = "Endocrine disrupter, neurotoxic, may cause long term adverse effects in the aquatic environment"
    @rd1272_symbols = ["Flammable gases", "Acute toxicity (oral, dermal, inhalation)", "Hazardous to the aquatic environment", "Danger"]
    @h_phrases = ["H224: Extremely flammable liquid and vapour", "H330: Fatal if inhaled", "H400: Very toxic to aquatic life", "H410: Very toxic to aquatic life with long lasting effects", "Flam. Liq. (Cat. 1): Flammable liquid", "Acute Tox. (Cat. 2 *): Acute toxicity", "Aquatic Acute (Cat. 1): Hazardous to the aquatic environment", "Aquatic Chronic (Cat. 1): Hazardous to the aquatic environment"]
  end

end
