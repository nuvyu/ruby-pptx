module PPTX
  module Shapes
    class FilledRectangle < Shape
      def initialize(transform, color)
        super(transform)
        @color = color
      end

      def base_xml
        # TODO replace cNvPr descr, id and name
        """
          <p:sp xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'
                xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>
              <p:nvSpPr>
                  <p:cNvPr id='2' name='Rectangle'/>
                  <p:cNvSpPr/>
                  <p:nvPr/>
              </p:nvSpPr>
              <p:spPr>
              </p:spPr>
          </p:sp>
        """
      end

      def build_node
        base_node.tap do |node|
          node.xpath('.//p:spPr', p: Presentation::NS).first.add_child build_solid_fill @color
        end
      end

      def build_solid_fill(rgb_color)
        fill_xml = """
        <a:solidFill xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>
          <a:srgbClr val='SETME'/>
        </a:solidFill>
        """

        Nokogiri::XML::DocumentFragment.parse(fill_xml).tap do |node|
          node.xpath('.//a:srgbClr', a: DRAWING_NS).first['val'] = rgb_color
        end
      end
    end
  end
end
