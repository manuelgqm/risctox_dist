require 'watir'
require_relative 'support/include_all_matcher'

browser = Watir::Browser.new :chrome, headless: true

RSpec.configure do |config|
  # only used on headed browser option
  config.before(:each) { @browser = browser }
  config.after(:suite) { browser.close unless browser.nil? }
end

class SpanElement
  attr_reader :label, :value

  def initialize(hash, browser)
    @label = browser.span(:id => hash + ".label").text
    @value = browser.span(:id => hash + ".value").text
  end

end

describe "'hydrogen cyanide' substance card" do
  before(:each) do
    @browser.goto("http://localhost:8081/en/dn_risctox_ficha_sustancia.asp?id_sustancia=953980")
  end

  describe "that has valid field labels and values" do
    it "should had a correct name" do
      name_element = SpanElement.new("name", @browser)
      expect(name_element.value).to include ('hydrogen cyanide')
      expect(name_element.label).to include ('Chemical name')
    end

    it "should had correct synonyms" do
      name_element = SpanElement.new("synonyms", @browser)
      expect(name_element.label).to include ('Synonyms')
      expect(name_element.value).to include ('hydrocyanic acid')
    end

    it "should had correct identification numbers" do
      identification_number_label_text = @browser.span(:id => "identification_numbers.label").text
      expect(identification_number_label_text).to include ('Identification numbers')
      name_element = SpanElement.new("cas_num", @browser)
      expect(name_element.label).to include ('CAS')
      expect(name_element.value).to include ('74-90-8')
      name_element = SpanElement.new("ec_einecs_num", @browser)
      expect(name_element.label).to include ('EC EINECS')
      expect(name_element.value).to include ('200-821-6')
    end
  end

end

describe "'ziram' substance card" do
  before(:each) do
    @browser.goto("http://localhost:8081/en/dn_risctox_ficha_sustancia.asp?id_sustancia=954057")
  end

  describe "that has valid field labels and values" do
    it "should had a correct name" do
      name_element = SpanElement.new("name", @browser)
      expect(name_element.label).to include ('Chemical name')
      expect(name_element.value).to include ('ziram')
    end

    it "should had correct synonyms" do
      name_element = SpanElement.new("synonyms", @browser)
      synonyms = [
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
      expect(name_element.value).to include_all synonyms
      expect(name_element.label).to include('Synonyms')
    end
  end

  it "should have correct trade names" do
    trade_name = SpanElement.new("trade_name", @browser)
    trade_names = [
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
    expect(trade_name.label).to include 'Trade name'
    expect(trade_name.value).to include_all trade_names
  end

  it "should had correct identification numbers" do
    identification_number_label_text = @browser.span(:id => "identification_numbers.label").text
    expect(identification_number_label_text).to include ('Identification numbers')
    name_element = SpanElement.new("cas_num", @browser)
    expect(name_element.value).to include ('137-30-4')
    expect(name_element.label).to include ('CAS')
    name_element = SpanElement.new("ec_einecs_num", @browser)
    expect(name_element.label).to include ('EC EINECS')
    expect(name_element.value).to include ('205-288-3')
  end

end
