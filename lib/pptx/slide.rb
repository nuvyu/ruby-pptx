require 'securerandom'

module PPTX
  class Slide < OPC::Part
    def initialize(package)
      part_name = package.next_part_with_prefix('ppt/slides/slide', '.xml')
      super(package, part_name)
      relationships.add('../slideLayouts/slideLayout7.xml', RELTYPE_SLIDE_LAYOUT)
    end

    def base_xml
      '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:cSld>
                <p:spTree>
                    <p:nvGrpSpPr>
                        <p:cNvPr id="1" name=""/>
                        <p:cNvGrpSpPr/>
                        <p:nvPr/>
                    </p:nvGrpSpPr>
                    <p:grpSpPr/>
                </p:spTree>
            </p:cSld>
            <p:clrMapOvr>
                <a:masterClrMapping/>
            </p:clrMapOvr>
        </p:sld>
        '''
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    end

    # Add rectangle filled with given color.
    # Give color in RGB hex format (e.g. '558ed5')
    def add_filled_rectangle(transform, color)
      shape = Shapes::FilledRectangle.new(transform, color)
      shape_tree_xml.add_child(shape.build_node)
    end

    def add_picture(transform, name, image)
      # TODO do something with name or remove
      rid = relationships.add(relative_part_name(image.part_name), PPTX::RELTYPE_IMAGE)

      shape = Shapes::Picture.new(transform, rid)
      shape_tree_xml.add_child(shape.build_node)
    end

    # Add a text box at the given position with the given dimensions
    #
    # Pass distance values in EMUs (use PPTX::CM)
    # Formatting values a are mostly raw OOXML Text Run Properties (rPr) element
    # attributes. Examples:
    # * sz: font size (use PPTX::POINT)
    # * b: 1 for bold
    # * i: 1 for italic
    # There is one custom property:
    # * color: aabbcc (hex RGB color)
    def add_textbox(transform, content, formatting={})
      shape = Shapes::Textbox.new(transform, content, formatting)
      shape_tree_xml.add_child(shape.build_node)
    end

    # Add a slide number that is recognized by Powerpoint as such and automatically updated
    # See textbox for formatting.
    def add_slide_number(transform, number, formatting={})
      shape = Shapes::SlideNumber.new(transform, number, formatting)
      shape_tree_xml.add_child(shape.build_node)
    end

    def shape_tree_xml
      @shape_tree_xml ||= doc.xpath('/p:sld/p:cSld/p:spTree').first
    end
  end
end
