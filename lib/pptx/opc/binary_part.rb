module PPTX
  module OPC
    class BinaryPart < OPC::BasePart
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

    class FilePart < BinaryPart
      def initialize(package, file)
        super(package, File.basename(file))
        @file = file
        @chunk_size = 16 * 1024
      end

      def marshal
        IO.read(@file)
      end

      def size
        File.size(@file)
      end

      def stream(out)
        File.open(@file, 'r') do |file|
          while chunk = file.read(@chunk_size)
            puts "Streaming file: #{chunk && chunk.bytesize} bytes"
            out << chunk
          end
        end
      end
    end

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
          puts "Streaming S3: #{chunk && chunk.bytesize} bytes"
          out << chunk
        end
      end
    end
  end
end

