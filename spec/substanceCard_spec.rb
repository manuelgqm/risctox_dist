require 'watir'
require_relative 'support/include_all_matcher'
require_relative 'fixture/ziram'
require_relative 'fixture/hydrogen_cyanide'
require_relative 'substance_card_page_object'

browser = Watir::Browser.new :chrome, headless: true

RSpec.configure do |config|
  config.before(:all) { @browser = browser }
  config.after(:suite) { browser.close unless browser.nil? } # only used on headed browser option
end

describe "'hydrogen cyanide' substance card" do
  before(:all) do
    @hydrogen_cyanide = Hydrogen_cyanide.new(953980)
    @page = PageObject.new(@browser, @hydrogen_cyanide.id)
    @page.go
  end

  describe "that has valid field labels and values" do
    it "should had a correct name" do
      expect(@page.name.label).to include ('Chemical name')
      expect(@page.name.value).to include @hydrogen_cyanide.name
    end

    it "must have valid synonyms" do
      expect(@page.synonyms.label).to include ('Synonyms')
      expect(@page.synonyms.value).to include_all @hydrogen_cyanide.synonyms
    end

    it "must have valid identification numbers" do
      identification_number_label_text = @browser.span(:id => "identification_numbers.label").text
      expect(identification_number_label_text).to include ('Identification numbers')
      expect(@page.cas_num.label).to include ('CAS')
      expect(@page.cas_num.value).to include @hydrogen_cyanide.cas_num
      expect(@page.ec_einecs_num.label).to include ('EC EINECS')
      expect(@page.ec_einecs_num.value).to include @hydrogen_cyanide.ce_einecs_num
    end
  end

  it "must have valid substance groups" do
    expect(@page.groups.label).to include 'Groups'
    expect(@page.groups.value).to include_all @hydrogen_cyanide.groups
  end

  it "must have valid substance uses" do
    expect(@page.uses.label).to include 'Uses'
    expect(@page.uses.value).to include_all @hydrogen_cyanide.uses
  end

  it "must have valid icsc numbers" do
    expect(@page.icsc_nums.label).to include 'International Chemical Safety Card (ICSC)'
    expect(@page.icsc_nums.value).to include_all @hydrogen_cyanide.icsc_nums
  end

  it "should had addition information" do
    additional_information_text = @browser.span(:id => "additional_information.label").text
    expect(additional_information_text).to include "Additional information"

    @page.toggle("secc-masinformacion")
    expect(@page.rd_num.label).to include "Index No"
    expect(@page.rd_num.value).to include @hydrogen_cyanide.rd_num
    expect(@page.molecular_formula.label).to include "Molecular formula"
    expect(@page.molecular_formula.value).to include @hydrogen_cyanide.molecular_formula

    @page.toggle("secc-concern_trade_union_list")
    expect(@page.concern_trade_union_reasons.label).to include "This substance is included in the List of Substances of concern for Trade Unions for the following reasons:"
    expect(@page.concern_trade_union_reasons.value).to include @hydrogen_cyanide.concern_trade_union_reasons
  end

end

describe "'ziram' substance card" do
  before(:all) do
    @ziram = Ziram.new(954057)
    @page = PageObject.new(@browser, @ziram.id)
    @page.go
  end

  describe "that has valid field labels and values" do
    it "should had a correct name" do
      expect(@page.name.label).to include ('Chemical name')
      expect(@page.name.value).to include @ziram.name
    end

    it "must have valid synonyms" do
      expect(@page.synonyms.label).to include('Synonyms')
      expect(@page.synonyms.value).to include_all @ziram.synonyms
    end
  end

  it "should have correct trade names" do
    expect(@page.trade_name.label).to include 'Trade name'
    expect(@page.trade_name.value).to include_all @ziram.trade_names
  end

  it "must have valid identification numbers" do
    identification_number_label_text = @browser.span(:id => "identification_numbers.label").text
    expect(identification_number_label_text).to include ('Identification numbers')
    expect(@page.cas_num.label).to include ('CAS')
    expect(@page.cas_num.value).to include @ziram.cas_num
    expect(@page.ec_einecs_num.label).to include ('EC EINECS')
    expect(@page.ec_einecs_num.value).to include @ziram.ec_einecs_num
  end

  it "must have valid substance groups" do
    expect(@page.groups.label).to include 'Groups'
    expect(@page.groups.value).to include_all @ziram.groups
  end

  it "must have valid substance uses" do
    expect(@page.uses.label).to include 'Uses'
    expect(@page.uses.value).to include_all @ziram.uses
  end

  it "must have valid icsc numbers" do
    expect(@page.icsc_nums.label).to include 'International Chemical Safety Card (ICSC)'
    expect(@page.icsc_nums.value).to include_all @ziram.icsc_nums
  end

  it "must have additional valid information" do
    additional_information_text = @browser.span(:id => "additional_information.label").text
    expect(additional_information_text).to include "Additional information"
    @page.toggle("secc-masinformacion")
    expect(@page.rd_num.label).to include "Index No"
    expect(@page.rd_num.value).to include @ziram.rd_num
    expect(@page.molecular_formula.label).to include "Molecular formula"
    expect(@page.molecular_formula.value).to include @ziram.molecular_formula

    @page.toggle("secc-concern_trade_union_list")
    expect(@page.concern_trade_union_reasons.label).to include "This substance is included in the List of Substances of concern for Trade Unions for the following reasons:"
    expect(@page.concern_trade_union_reasons.value).to include @ziram.concern_trade_union_reasons
  end

end
