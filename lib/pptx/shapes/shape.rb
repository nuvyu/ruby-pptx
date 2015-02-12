module PPTX
  module Shapes
    class Shape
      def initialize(transform)
        @transform = transform
      end

      def base_node
        @base_node ||= Nokogiri::XML::DocumentFragment.parse(base_xml).tap do |node|
          set_shape_properties(node, *@transform)
        end
      end

      def set_shape_properties(node, x, y, width, height)
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
    end
  end
end