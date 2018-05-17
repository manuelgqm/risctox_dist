class PageObject

  def initialize(browser, substance_id)
    @browser = browser
    @substance_id = substance_id
  end

  def name
    field_element("name")
  end

  def rd_num
    field_element("rd_num")
  end

  def molecular_formula
    field_element("molecular_formula")
  end

  def concern_trade_union_reasons
    field_element("concern_trade_union_reasons")
  end

  def synonyms
    field_element("synonyms")
  end

  def cas_num
    field_element("cas_num")
  end

  def ec_einecs_num
    field_element("ec_einecs_num")
  end

  def groups
    field_element("groups")
  end

  def uses
    field_element("uses")
  end

  def icsc_nums
    field_element("icsc_nums")
  end

  def trade_name
    field_element("trade_name")
  end

  def rd1272_symbols
    element("rd1272_symbols")
  end

  def H_phrases
    element("H_phrases")
  end

  def rd1272_labeling
    element("rd1272_labeling")
  end

  def rd1272_notes
    element("rd1272_notes")
  end

  def carcinogen_rd1272
    field_element("carcinogen_rd1272")
  end

  def carcinogen_iarc
    element("carcinogen_iarc")
  end

  def carcinogen_iarc_group
    element("carcinogen_iarc_group")
  end

  def carcinogen_iarc_volume
    element("carcinogen_iarc_volume")
  end

  def carcinogen_iarc_notes
    element("carcinogen_iarc_notes")
  end

  def toggle(element_id)
    script = "arguments[0].setAttribute('style', 'display:block')"
    element = @browser.element(:id => element_id)
    @browser.execute_script(script, element)
  end

  def go
    @browser.goto("http://localhost:8081/en/dn_risctox_ficha_sustancia_new.asp?id_sustancia=" + @substance_id.to_s)
  end

  private

  def field_element(id)
    return PageField.new(id, @browser)
  end

  def element(id)
    return PageElement.new(id, @browser)
  end
end

class PageField
  attr_reader :label, :value

  def initialize(hash, browser)
    @label = browser.element(:id => hash + ".label").text
    @value = browser.element(:id => hash + ".value").text
  end
end

class PageElement
  attr_reader :text

  def initialize(id, browser)
    @text = browser.element(:id => id).text
  end
end
