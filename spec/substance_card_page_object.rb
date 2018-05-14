class PageElement
  attr_reader :label, :value

  def initialize(hash, browser)
    @label = browser.element(:id => hash + ".label").text
    @value = browser.element(:id => hash + ".value").text
  end

end

class PageObject

  def initialize(browser, substance_id)
    @browser = browser
    @substance_id = substance_id
  end

  def name
    element("name")
  end

  def rd_num
    element("rd_num")
  end

  def molecular_formula
    element("molecular_formula")
  end

  def concern_trade_union_reasons
    element("concern_trade_union_reasons")
  end

  def synonyms
    element("synonyms")
  end

  def cas_num
    element("cas_num")
  end

  def ec_einecs_num
    element("ec_einecs_num")
  end

  def groups
    element("groups")
  end

  def uses
    element("uses")
  end

  def icsc_nums
    element("icsc_nums")
  end

  def trade_name
    element("trade_name")
  end

  def toggle(element_id)
    script = "arguments[0].setAttribute('style', 'display:block')"
    element = @browser.element(:id => element_id)
    @browser.execute_script(script, element)
  end

  def go
    @browser.goto("http://localhost:8081/en/dn_risctox_ficha_sustancia.asp?id_sustancia=" + @substance_id.to_s)
  end

  private
  def element(id)
    return PageElement.new(id, @browser)
  end
end
