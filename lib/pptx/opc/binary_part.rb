module PPTX
  module OPC
    class BinaryPart < BasePart
      def initialize(package, original_name)
        name = "ppt/media/binary-#{SecureRandom.hex(10)}#{File.extname(original_name)}"
        super(package, name)
        @original_name = original_name
      end

      def doc
        throw "Binary parts can't be parsed as XML"
      end

      def marshal
        throw 'Implement me!'
      end
    end
  end
end

