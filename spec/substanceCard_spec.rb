require 'watir'

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
      expect(name_element.label).to include ('Synonyms')
      expect(name_element.value).to include ('Zinc dimethyldithiocarbamate')
      expect(name_element.value).to include ('Methyl zimate')
      expect(name_element.value).to include ('Dimethyldithiocarbamic acid zinc salt')
      expect(name_element.value).to include ('Bis (dimethylcarbamodithioato - S, S`) zinc')
      expect(name_element.value).to include ('Bis (dimethyldithiocarbamato) zinc')
      expect(name_element.value).to include ('Zinc bis (dimethylthiocarbamoyl) disulfide')
      expect(name_element.value).to include ('Carbazinc')
      expect(name_element.value).to include ('Fungostop')
      expect(name_element.value).to include ('Zinc bis (dimethyldithiocarbamoyl) disulfide')
      expect(name_element.value).to include ('Methyl Ziram')
      expect(name_element.value).to include ('Dimethylcarbamodithioic acid, zinc complex')
      expect(name_element.value).to include ('Dimethyldithiocarbamate zinc salt')
      expect(name_element.value).to include ('Zinc N, N - dimethyldithiocarbamate')
      expect(name_element.value).to include ('Amylzimate')
      expect(name_element.value).to include ('Carbamic acid, dimethyldithio - , zinc salt (2:1)')
      expect(name_element.value).to include ('Ciram')
      expect(name_element.value).to include ('Methyl zineb')
      expect(name_element.value).to include ('Bis (dimethylcarbamodithiato - S, S`) - zinc')
      expect(name_element.value).to include ('(SP - 4 - 1) - Bis (dimethylcarbamodithiato - S, S`) - zinc')
      expect(name_element.value).to include ('(T - 4) - Bis (dimethylcarbamodithioato - S, S`) - zinc')
      expect(name_element.value).to include ('(T - 4) - Bis (dimethyldithiocarbamato - S, S` ) zinc')
      expect(name_element.value).to include ('Methyl cymate')
    end
  end

end
