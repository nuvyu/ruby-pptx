module PPTX
  module Shapes
    class Shape
      def initialize(transform, shape_formatting={})
        @transform = transform
        @shape_formatting = shape_formatting
      end

      def base_node
        @base_node ||= Nokogiri::XML::DocumentFragment.parse(base_xml).tap do |node|
          build_shape_properties(node, *@transform)
          set_shape_properties(node, @shape_formatting)
        end
      end

      def build_shape_properties(node, x, y, width, height)
        xml = """
          <p:spPr xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'
                  xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>
              <a:xfrm>
                  <a:off x='#{x.to_i}' y='#{y.to_i}'/>
                  <a:ext cx='#{width.to_i}' cy='#{height.to_i}'/>
              </a:xfrm>
              <a:prstGeom prst='rect'>
                  <a:avLst/>
              </a:prstGeom>
          </p:spPr>
          """

        node.xpath('.//p:spPr', p: Presentation::NS)
            .first
            .replace(Nokogiri::XML::DocumentFragment.parse(xml))
      end


      def set_shape_properties(node, formatting)
        bg = formatting.delete(:bg)
        node.xpath('.//p:spPr', p: Presentation::NS).first.add_child build_solid_fill(bg) if(bg)
      end

      def extract_shape_properties(formatting)
        shape_formats = [:bg]

        shape_formatting = {}
        formatting.each do |type, value|
          if(shape_formats.include?(type))
            shape_formatting[type] = value
            formatting.delete(type)
          end
        end
        shape_formatting
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