require 'securerandom'

module PPTX
  class Slide < OPC::Part
    def initialize(package)
      part_name = package.next_part_with_prefix('ppt/slides/slide', '.xml')
      super(package, part_name, 'ppt/slides/slide1.xml')
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    end

    def add_picture(transform, name, image)
      image_name = "ppt/media/image-#{SecureRandom.hex(10)}#{File.extname(name)}"
      @package.set_part(image_name, image.read)
      rid = relationships.add(relative_part_name(image_name), PPTX::RELTYPE_IMAGE)

      shape = Shapes::Picture.new(transform, rid)
      shape_tree_xml.add_child(shape.build_node)
    end

    # Add a text box at the given position with the given dimensions
    #
    # Pass distance values in EMUs (use PPTX::CM)
    # Formatting values a are raw OOXML Text Run Properties (rPr) element
    # attributes. Examples:
    # * sz: font size (use PPTX::POINT)
    # * b: 1 for bold
    # * i: 1 for italic
    def add_textbox(transform, content, formatting={})
      shape = Shapes::Textbox.new(transform, content, formatting)
      shape_tree_xml.add_child(shape.build_node)
    end

    def shape_tree_xml
      @shape_tree_xml ||= doc.xpath('/p:sld/p:cSld/p:spTree').first
    end
  end
end
