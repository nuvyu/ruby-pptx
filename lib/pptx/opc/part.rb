module PPTX
  module OPC
    # A part without relationships.
    # * template_part allows defining a part name that can be used as initial XML document
    #   when a part with the name doesn't exist yet, but the part shouldn't be created
    #   with an empty XML document (e.g. when the creation of the whole document isn't
    #   implemented, but only certain operations).
    class BasePart
      def initialize(package, part_name, template_part = nil)
        @package = package
        @part_name = part_name
        if template_part
          @doc = @package.part_xml(template_part)
        end
        package.set_part(part_name, self, content_type)
      end

      def content_type
        nil
      end

      def doc
        @doc ||= @package.base_part_xml(@part_name)
      end

      def marshal
        doc.to_s
      end

      def part_name
        @part_name
      end
    end

    class Part < BasePart
      def initialize(package, part_name, template_part = nil)
        super(package, part_name, template_part)
        rel_template_part = template_part ? relationship_file(template_part) : nil
        @relationships = Relationships.new(package, relationship_file, rel_template_part)
      end

      def relative_part_name(name)
        name[File.dirname(part_name).size + 1..-1]
      end

      def relationships
        @relationships
      end

      def relationship_file(part_name = nil)
        part_name ||= @part_name
        File.join(File.dirname(part_name), '_rels', File.basename(part_name) + '.rels')
      end
    end
  end
end
