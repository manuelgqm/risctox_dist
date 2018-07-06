class Ziram
  attr_reader :id
  attr_reader :name, :synonyms, :trade_names, :cas_num, :ec_einecs_num, :groups
  attr_reader :uses, :icsc_nums, :rd_num, :molecular_formula, :concern_trade_union_reasons
  attr_reader :rd1272_symbols, :H_phrases, :rd1272_labeling, :regulations

  def initialize()
    @id = 954057
    @name = "ziram"
    @synonyms = [
      'Zinc dimethyldithiocarbamate',
      'Methyl zimate',
      'Dimethyldithiocarbamic acid zinc salt',
      'Bis (dimethylcarbamodithioato - S, S`) zinc',
      'Bis (dimethyldithiocarbamato) zinc',
      'Zinc bis (dimethylthiocarbamoyl) disulfide',
      'Carbazinc',
      'Fungostop',
      'Zinc bis (dimethyldithiocarbamoyl) disulfide',
      'Methyl Ziram',
      'Dimethylcarbamodithioic acid, zinc complex',
      'Dimethyldithiocarbamate zinc salt',
      'Zinc N, N - dimethyldithiocarbamate',
      'Amylzimate',
      'Carbamic acid, dimethyldithio - , zinc salt (2:1)',
      'Ciram',
      'Methyl zineb',
      'Bis (dimethylcarbamodithiato - S, S`) - zinc',
      '(SP - 4 - 1) - Bis (dimethylcarbamodithiato - S, S`) - zinc',
      '(T - 4) - Bis (dimethylcarbamodithioato - S, S`) - zinc',
      '(T - 4) - Bis (dimethyldithiocarbamato - S, S` ) zinc',
      'Methyl cymate'
    ]
    @trade_names = [
      "Aaprotect",
      "Aavolex",
      "Aazira",
      "Accelerator L",
      "Aceto ZDED",
      "Aceto ZDMD",
      "Alcobam ZM",
      "Antene",
      "Corona Corozate",
      "Corozate",
      "Cuman",
      "Cuman L",
      "Cymate",
      "Drupina 90",
      "Eptac 1",
      "Fuclasin",
      "Fuclasin Ultra",
      "Fuklasin",
      "Hermat ZDM",
      "Hexazir",
      "KarbamWhite",
      "Methasan",
      "Methazate",
      "Mezene",
      "Milbam",
      "Molurame",
      "Mycronil",
      "Pomarsol Z - forte",
      "Prodaram",
      "Rhodiacid",
      "Soxinal PZ",
      "Soxinol PZ",
      "Tricarbamix Z",
      "Triscabol",
      "Tsimat",
      "Vancide",
      "Vancide MZ - 96",
      "Vulcacure",
      "Vulcacure ZM",
      "Vulkacite L",
      "Z 75",
      "Zarlate",
      "ZC",
      "Z - C Spray",
      "Zerlate",
      "Zimate",
      "Zincmate",
      "Ziram F4",
      "Ziram W76",
      "Ziramvis",
      "Zirasan",
      "Zirasan 90",
      "Zirberk",
      "Zirex 90",
      "Ziride",
      "Zirthane",
      "Zitox"
    ]
    @cas_num = "137-30-4"
    @ec_einecs_num = "205-288-3"
    @groups = ["organo phosphorus and carbamates"]
    @uses = ["fungicide", "pesticide"]
    @icsc_nums = ["0348"]
    @icsc_nums_links = ["http://www.ilo.org/dyn/icsc/showcard.display?p_lang=en&p_card_id=0348"]
    @rd_num = "006-012-00-2"
    @molecular_formula = "C6H12N2S4Zn"
    @concern_trade_union_reasons = "Endocrine disrupter, neurotoxic, sensitizer, may cause long term adverse effects in the aquatic environment"
    @rd1272_symbols = ["Acute toxicity (oral, dermal, inhalation)", "Respiratory sensitisation", "Corrosive to metals", "Hazardous to the aquatic environment", "Danger"]
    @H_phrases = ["H330: Fatal if inhaled", "Acute Tox. (Cat. 2 *): Acute toxicity", "H302: Harmful if swallowed", "Acute Tox. (Cat. 4 *): Acute toxicity", "H373 **: May cause damage to organs through prolonged or repeated exposure", "STOT RE (Cat. 2 *): Specific target organ toxicity - repeated exposure", "H335: May cause respiratory irritation", "STOT SE (Cat. 3): Specific target organ toxicity - single exposure", "H318: Causes serious eye damage", "Eye Dam. (Cat. 1): Serious eye damage/eye irritation", "H317: May cause an allergic skin reaction", "Skin Sens. (Cat. 1): Respiratory/skin sensitization", "H400: Very toxic to aquatic life", "Aquatic Acute (Cat. 1): Hazardous to the aquatic environment", "H410: Very toxic to aquatic life with long lasting effects", "Aquatic Chronic (Cat. 1): Hazardous to the aquatic environment"]
    @rd1272_labeling = ["Factor M = 100"]
    @regulations = [
      'Banned biocide',
      'Authorised pesticide',
      'CoRAP evaluation'
    ]
  end

end
