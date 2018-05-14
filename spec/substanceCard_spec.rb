require 'watir'
require_relative 'support/include_all_matcher'
require_relative 'fixture/ziram'
require_relative 'fixture/hydrogen_cyanide'

browser = Watir::Browser.new :chrome, headless: true

RSpec.configure do |config|
  # only used on headed browser option
  config.before(:all) { @browser = browser }
  config.after(:suite) { browser.close unless browser.nil? }
end

class SpanElement
  attr_reader :label, :value

  def initialize(hash, browser)
    @label = browser.span(:id => hash + ".label").text
    @value = browser.span(:id => hash + ".value").text
  end

end

class PageObject

  def initialize(browser)
    @browser = browser
  end

  def toggle(element_id)
    script = "arguments[0].setAttribute('style', 'display:block')"
    element = @browser.td(:id => element_id)
    @browser.execute_script(script, element)
  end
end

describe "'hydrogen cyanide' substance card" do
  before(:all) do
    @hydrogen_cyanide = Hydrogen_cyanide.new(@browser)
    @hydrogen_cyanide.go
  end

  describe "that has valid field labels and values" do
    it "should had a correct name" do
      name_element = SpanElement.new("name", @browser)
      expect(name_element.label).to include ('Chemical name')
      expect(name_element.value).to include @hydrogen_cyanide.name
    end

    it "should had correct synonyms" do
      name_element = SpanElement.new("synonyms", @browser)
      expect(name_element.label).to include ('Synonyms')
      expect(name_element.value).to include_all @hydrogen_cyanide.synonyms
    end

    it "should had correct identification numbers" do
      identification_number_label_text = @browser.span(:id => "identification_numbers.label").text
      expect(identification_number_label_text).to include ('Identification numbers')
      name_element = SpanElement.new("cas_num", @browser)
      expect(name_element.label).to include ('CAS')
      expect(name_element.value).to include @hydrogen_cyanide.cas_num
      name_element = SpanElement.new("ec_einecs_num", @browser)
      expect(name_element.label).to include ('EC EINECS')
      expect(name_element.value).to include @hydrogen_cyanide.ce_einecs_num
    end
  end

  it "should had correct substance groups" do
    groups_element = SpanElement.new("groups", @browser)
    expect(groups_element.label).to include 'Groups'
    expect(groups_element.value).to include_all @hydrogen_cyanide.groups
  end

  it "should had correct substance uses" do
    groups_element = SpanElement.new("uses", @browser)
    expect(groups_element.label).to include 'Uses'
    expect(groups_element.value).to include_all @hydrogen_cyanide.uses
  end

  it "should had correct icsc numbers" do
    element = SpanElement.new("icsc_nums", @browser)
    expect(element.label).to include 'International Chemical Safety Card (ICSC)'
    expect(element.value).to include_all @hydrogen_cyanide.icsc_nums
  end

  it "should had addition information" do
    additional_information_text = @browser.span(:id => "additional_information.label").text
    expect(additional_information_text).to include "Additional information"
    page = PageObject.new(@browser)
    page.toggle("secc-masinformacion")
    element = SpanElement.new("rd_num", @browser)
    expect(element.label).to include "Index No"
    expect(element.value).to include @hydrogen_cyanide.rd_num
    element = SpanElement.new("molecular_formula", @browser)
    expect(element.label).to include "Molecular formula"
    expect(element.value).to include @hydrogen_cyanide.molecular_formula
  end

end

describe "'ziram' substance card" do
  before(:all) do
    @ziram = Ziram.new(@browser)
    @ziram.go
  end

  describe "that has valid field labels and values" do
    it "should had a correct name" do
      name_element = SpanElement.new("name", @browser)
      expect(name_element.label).to include ('Chemical name')
      expect(name_element.value).to include @ziram.name
    end

    it "should had correct synonyms" do
      name_element = SpanElement.new("synonyms", @browser)
      expect(name_element.label).to include('Synonyms')
      expect(name_element.value).to include_all @ziram.synonyms
    end
  end

  it "should have correct trade names" do
    trade_name = SpanElement.new("trade_name", @browser)

    expect(trade_name.label).to include 'Trade name'
    expect(trade_name.value).to include_all @ziram.trade_names
  end

  it "should had correct identification numbers" do
    identification_number_label_text = @browser.span(:id => "identification_numbers.label").text
    expect(identification_number_label_text).to include ('Identification numbers')
    name_element = SpanElement.new("cas_num", @browser)
    expect(name_element.label).to include ('CAS')
    expect(name_element.value).to include @ziram.cas_num
    name_element = SpanElement.new("ec_einecs_num", @browser)
    expect(name_element.label).to include ('EC EINECS')
    expect(name_element.value).to include @ziram.ec_einecs_num
  end

  it "should had correct substance groups" do
    groups_element = SpanElement.new("groups", @browser)
    expect(groups_element.label).to include 'Groups'
    expect(groups_element.value).to include_all @ziram.groups
  end

  it "should had correct substance uses" do
    element = SpanElement.new("uses", @browser)
    expect(element.label).to include 'Uses'
    expect(element.value).to include_all @ziram.uses
  end

  it "should had correct icsc numbers" do
    element = SpanElement.new("icsc_nums", @browser)
    expect(element.label).to include 'International Chemical Safety Card (ICSC)'
    expect(element.value).to include_all @ziram.icsc_nums
  end

  it "should had addition information" do
    additional_information_text = @browser.span(:id => "additional_information.label").text
    expect(additional_information_text).to include "Additional information"
    page = PageObject.new(@browser)
    page.toggle("secc-masinformacion")
    element = SpanElement.new("rd_num", @browser)
    expect(element.label).to include "Index No"
    expect(element.value).to include @ziram.rd_num
    element = SpanElement.new("molecular_formula", @browser)
    expect(element.label).to include "Molecular formula"
    expect(element.value).to include @ziram.molecular_formula
  end

end
