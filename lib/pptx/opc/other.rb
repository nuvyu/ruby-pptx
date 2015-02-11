require 'securerandom'

module PPTX
  module OPC

    class ContentTypes < BasePart
      def set_override(part_name, value)
        type_list = doc.xpath('c:Types', c: CONTENTTYPE_NS).first
        override = type_list.xpath("./xmlns:Override[@PartName='#{part_name}']").first
        unless override
          override = Nokogiri::XML::Node.new('Override', doc)
          override['PartName'] = '/' + part_name
          type_list.add_child(override)
        end
        override['ContentType'] = value
      end
    end

    class Relationships < BasePart
      def add(relative_part_name, type)
        ref_id = "rId#{SecureRandom.hex(10)}"

        relationship = Nokogiri::XML::Node.new('Relationship', doc)
        relationship['Id'] = ref_id
        relationship['Target'] = relative_part_name
        relationship['Type'] = type
        list_xml.add_child(relationship)

        ref_id
      end

      def list_xml
        @list_xml ||= doc.xpath('r:Relationships', r: OPC::RELATIONSHIP_NS).first
      end
    end
  end
end
