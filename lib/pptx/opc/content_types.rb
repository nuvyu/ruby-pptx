module PPTX
  module OPC
    class ContentTypes < BasePart
      NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

      def set_override(part_name, value)
        type_list = doc.xpath('c:Types', c: NS).first
        override = type_list.xpath("./xmlns:Override[@PartName='#{part_name}']").first
        unless override
          override = Nokogiri::XML::Node.new('Override', doc)
          override['PartName'] = '/' + part_name
          type_list.add_child(override)
        end
        override['ContentType'] = value
      end
    end
  end
end
