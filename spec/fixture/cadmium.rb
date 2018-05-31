class Cadmium
  attr_reader :id, :carcinogen_rd1272, :carcinogen_iarc, :carcinogen_iarc_group, :carcinogen_iarc_volume, :carcinogen_other_sources_categories, :carcinogen_other_sources_definitions
  def initialize()
    @id = 955314
    @carcinogen_rd1272 = ["Carcinogen level: 1B"]
    @carcinogen_iarc_group = "1"
    @carcinogen_iarc_volume = "58, 100C; 2012"
    @carcinogen_other_sources_categories = ["G-A2"]
    @carcinogen_other_sources_definitions = ["Suspected human carcinogen"]
  end
end
