require_relative 'pptx/opc'
require_relative 'pptx/presentation'
require_relative 'pptx/slide'
require_relative 'pptx/shapes'

# Assumptions
# * XML namespace prefixes remain the same as in template
# * Main presentation is in ppt/presentation.xml
# * Base presentation is trusted. Nothing is done to prevent directory traversal.
# If you wanted to open arbitrary presentation, you'd have to make sure to no longer make these.
module PPTX
  DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  DOC_RELATIONSHIP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  RELTYPE_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
  RELTYPE_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
  RELTYPE_SLIDE_LAYOUT =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
  CM = 360000  # 1 centimeter in OpenXML EMUs
  POINT = 100  # font size point
end
