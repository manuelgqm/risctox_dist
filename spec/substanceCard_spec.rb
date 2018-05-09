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
  end

end
