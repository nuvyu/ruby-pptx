require 'securerandom'

module PPTX
  class Slide < OPC::Part
    def initialize(package, part_name)
      super(package, part_name, 'ppt/slides/slide1.xml')
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    end

    def add_picture(transform, name, image)
      # TODO replace cNvPr descr, id and name
      picture_xml = """
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
                <a:prstGeom prst='rect'>
                    <a:avLst/>
                </a:prstGeom>
            </p:spPr>
        </p:pic>
      """

      # add transform to shape properties (spPr)
      node = Nokogiri::XML::DocumentFragment.parse(picture_xml)
      node.xpath('.//p:spPr', p: Presentation::NS)
          .first
          .prepend_child(create_transform_xml(*transform))

      image_name = "ppt/media/image-#{SecureRandom.hex(10)}#{File.extname(name)}"
      @package.set_part(image_name, image.read)
      rid = relationships.add(relative_part_name(image_name), PPTX::RELTYPE_IMAGE)

      node.xpath('.//a:blip', a: DRAWING_NS).first['r:embed'] = rid

      shape_tree_xml.add_child(node)
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
      # TODO replace cNvPr descr, id and name
      shape_xml = """
                <p:sp xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'
                      xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>
                    <p:nvSpPr>
                        <p:cNvPr id='2' name='TextBox'/>
                        <p:cNvSpPr txBox='1'/>
                        <p:nvPr/>
                    </p:nvSpPr>
                    <p:spPr>
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
      node.xpath('.//p:spPr', p: Presentation::NS)
          .first
          .prepend_child(create_transform_xml(*transform))

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
        par_node.xpath('./a:p/a:r/a:rPr', a: DRAWING_NS).first.tap do |rpr|
          formatting.each do |key, val|
            rpr[key] = val
          end
        end
        txbody_node.add_child par_node
      end

      shape_tree_xml.add_child(node)
    end

    def shape_tree_xml
      @shape_tree_xml ||= doc.xpath('/p:sld/p:cSld/p:spTree').first
    end

    def create_transform_xml(x, y, width, height)
      Nokogiri::XML::DocumentFragment.parse("""
        <a:xfrm xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>
            <a:off x='#{x.to_i}' y='#{y.to_i}'/>
            <a:ext cx='#{width.to_i}' cy='#{height.to_i}'/>
        </a:xfrm>
        """)
    end
  end
end
