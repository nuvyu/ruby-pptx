require 'pathname'

module PPTX
  module OPC
    class Part < BasePart
      def initialize(package, part_name)
        super(package, part_name)
        @relationships = Relationships.new(package, relationship_part_name)
      end

      def relative_part_name(name)
        source = Pathname.new(File.dirname(part_name))
        target = Pathname.new(name)
        target.relative_path_from(source).to_s
      end

      def relationships
        @relationships
      end

      def relationship_part_name(part_name = nil)
        part_name ||= @part_name
        File.join(File.dirname(part_name), '_rels', File.basename(part_name) + '.rels')
      end
    end
  end
end
