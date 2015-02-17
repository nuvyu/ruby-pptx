module PPTX
  module OPC
    # A part without relationships.
    class BasePart
      def initialize(package, part_name)
        @package = package
        @part_name = part_name
        package.set_part(part_name, self, content_type)
      end

      def base_xml
        nil
      end

      def content_type
        nil
      end

      def doc
        @doc ||= Nokogiri::XML(template || base_xml)
      end

      def marshal
        doc.to_s
      end

      def part_name
        @part_name
      end

      def template
        @package.template_part(@part_name)
      end
    end
  end
end
