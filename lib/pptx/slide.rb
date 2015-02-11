module PPTX
  class Slide < OPC::Part
    def initialize(package, part_name)
      super(package, part_name, 'ppt/slides/slide1.xml')
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    end

    # Pass distance values in EMUs (use PPTX::CM)
    def add_textbox(x, y, width, height, content, formatting=nil)
      x, y, width, height = x.to_i, y.to_i, width.to_i, height.to_i

      # TODO Find out whether some of these ids have to be fixed, e.g. CNvPr id
      shape_xml = """
                <p:sp xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'
                      xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>
                    <p:nvSpPr>
                        <p:cNvPr id='2' name='TextBox'/>
                        <p:cNvSpPr txBox='1'/>
                        <p:nvPr/>
                    </p:nvSpPr>
                    <p:spPr>
                        <a:xfrm>
                            <a:off x='#{x}' y='#{y}'/>
                            <a:ext cx='#{width}' cy='#{height}'/>
                        </a:xfrm>
                        <a:prstGeom prst='rect'>
                            <a:avLst/>
                        </a:prstGeom>
                        <a:noFill/>
                    </p:spPr>
                    <p:txBody>
                        <a:bodyPr rtlCol='0' wrap='none'>
                            <a:spAutoFit/>
                        </a:bodyPr>
                        <a:lstStyle/>
                    </p:txBody>
                </p:sp>"""
      node = Nokogiri::XML::DocumentFragment.parse(shape_xml)

      paragraph_xml = """
        <a:p xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>
            <a:r>
                <a:rPr lang='en-US' smtClean='0'/>
                <a:t></a:t>
            </a:r>
            <a:endParaRPr lang='en-US'/>
        </a:p>
      """

      txbody_node = node.xpath('./p:sp/p:txBody', p:Presentation::NS).first

      content.split("\n").each do |line|
        par_node = Nokogiri::XML::DocumentFragment.parse(paragraph_xml)
        par_node.xpath('./a:p/a:r/a:t', a: DRAWING_NS).first.tap { |tn| tn.content = line }
        txbody_node.add_child par_node
      end

      shape_tree_xml.add_child(node)
    end

    def shape_tree_xml
      @shape_tree_xml ||= doc.xpath('/p:sld/p:cSld/p:spTree').first
    end
  end
end
