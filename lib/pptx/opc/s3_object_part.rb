module PPTX
  module OPC
    class S3ObjectPart < BinaryPart
      def initialize(package, s3obj, size)
        super(package, File.basename(s3obj.key))
        @object = s3obj
        @size = size
      end

      def marshal
        @object.read
      end

      def size
        @size
      end

      def stream(out)
        @object.read do |chunk|
          out << chunk
        end
      end
    end
  end
end
