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
  DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  DOC_RELATIONSHIP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  RELTYPE_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
  CM = 360000  # 1 centimeter in OpenXML EMUs
  POINT = 100  # font size point
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

  text = "Text box text\n:)\nwith<stuff>&to<be>escaped\nYay!"

  tree = slide.shape_tree_xml

  # Remove existing pictures
  tree.xpath('./p:pic').each do |pic|
    pic.unlink
  end

  slide.add_textbox(14*PPTX::CM, 6*PPTX::CM, 10*PPTX::CM, 10*PPTX::CM, text)
  slide.add_textbox(2*PPTX::CM, 1*PPTX::CM, 22*PPTX::CM, 3*PPTX::CM,
                    'Title :)', sz: 45*PPTX::POINT)

  puts tree.to_s

  presentation = pkg.presentation
  presentation.add_slide(slide)

  File.open('tmp/generated.pptx', 'wb') {|f| f.write(pkg.to_zip) }
end

main
