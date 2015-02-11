module PPTX
  module OPC
    RELATIONSHIP_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    CONTENTTYPE_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
  end
end

require_relative 'opc/package'
require_relative 'opc/part'
require_relative 'opc/other'
