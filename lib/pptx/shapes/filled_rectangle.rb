module PPTX
  module Shapes
    class FilledRectangle < Shape
      def initialize(transform, formatting={})
        super(transform, extract_shape_properties(formatting))
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
        base_node
      end
    end
  end
end
