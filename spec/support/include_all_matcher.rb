RSpec::Matchers.define :include_all do |include_items|
  match do |given|
    @errors = include_items.reject { |item| given.include?(item) }
    @errors.empty?
  end

  failure_message do |given|
    "did not include \"#{@errors.join('\", \"')}\""
  end

  failure_message_when_negated do |given|
     "everything was included"
  end

  description do |given|
    "includes all of #{include_items.join(', ')}"
  end
end
