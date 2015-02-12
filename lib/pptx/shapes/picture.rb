module PPTX
  module Shapes
    class Picture < Shape
      def initialize(transform, relationship_id)
        super(transform)
        @relationship_id = relationship_id
      end

      def base_xml
        # TODO replace cNvPr descr, id and name
        """
          <p:pic xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'
                 xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>
            <p:nvPicPr>
                <p:cNvPr descr='test_photo.jpg' id='2' name='Picture 1'/>
                <p:cNvPicPr>
                    <a:picLocks noChangeAspect='1'/>
                </p:cNvPicPr>
                <p:nvPr/>
            </p:nvPicPr>
            <p:blipFill>
                <a:blip r:embed='REPLACEME'/>
                <a:stretch>
                    <a:fillRect/>
                </a:stretch>
            </p:blipFill>
            <p:spPr>
            </p:spPr>
        </p:pic>
        """
      end

      def build_node
        base_node.tap do |node|
          node.xpath('.//a:blip', a: DRAWING_NS).first['r:embed'] = @relationship_id
        end
      end
    end
  end
end
