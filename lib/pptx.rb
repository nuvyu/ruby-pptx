require 'bundler'

# Bundler setup instead of require and doing manual requires significantly speeds up startup
# as not the whole project's dependencies are loaded.
Bundler.setup(:default, 'development')

require_relative 'pptx/opc'
require_relative 'pptx/presentation'
require_relative 'pptx/slide'

# Assumptions
# * XML namespace prefixes remain the same as in template
# * Main presentation is in ppt/presentation.xml
# * Base presentation is trusted. Nothing is done to prevent directory traversal.
# If you wanted to open arbitrary presentation, you'd have to make sure to no longer make these.
module PPTX
  DOC_RELATIONSHIP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  RELTYPE_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
end


def main
  pkg = PPTX::OPC::Package.new

  # TODO move this to presentation
  # app = pkg.part_xml 'docProps/app.xml'
  # slides = app.xpath('/ep:Properties/ep:Slides',
  #   ep: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties').first
  # slides.content = 1
  # pkg.parts['docProps/app.xml'] = app.to_s

  slide = PPTX::Slide.new(pkg, 'ppt/slides/slide2.xml')

  presentation = pkg.presentation
  presentation.add_slide(slide)

  File.open('tmp/generated.pptx', 'wb') {|f| f.write(pkg.to_zip) }
end

main
