require 'watir'

class Ziram
  attr_reader :name, :synonyms, :trade_names, :cas_num, :ec_einecs_num, :groups
  attr_reader :uses, :icsc_nums, :rd_num, :molecular_formula, :concern_trade_union_reasons
  def initialize(browser)
    @browser = browser
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
  end

  def go
    @browser.goto("http://localhost:8081/en/dn_risctox_ficha_sustancia.asp?id_sustancia=954057")
  end
end
