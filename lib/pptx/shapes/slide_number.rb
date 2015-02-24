module PPTX
  module Shapes
    class SlideNumber < Textbox
      def initialize(transform, slidenum, formatting={})
        super(transform, slidenum.to_s, formatting)
      end

      def build_node
        node = super

        placeholder_xml = """
          <p:ph xmlns:p='#{Presentation::NS}' sz='quarter' type='sldNum'/>"""
        placeholder = Nokogiri::XML::DocumentFragment.parse(placeholder_xml)
        node.xpath('./p:sp/p:nvSpPr/p:nvPr', p:Presentation::NS).first.add_child placeholder

        node
      end

      def build_paragraph(slidenum)
        paragraph_xml = """
          <a:p xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>
              <a:fld id='REPLACE ME' type='slidenum'>
                  <a:rPr lang='en-US' smtClean='0'/>
                  <a:t>REPLACE ME</a:t>
              </a:fld>
              <a:endParaRPr lang='en-US'/>
          </a:p>
        """

        Nokogiri::XML::DocumentFragment.parse(paragraph_xml).tap do |node|
          node.xpath('./a:p/a:fld/a:t', a: DRAWING_NS).first.content = slidenum
          field_id = "{#{SecureRandom.uuid}}"
          node.xpath('./a:p/a:fld', a: DRAWING_NS).first['id'] = field_id
        end
      end
    end
  end
end
