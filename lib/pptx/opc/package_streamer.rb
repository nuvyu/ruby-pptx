module PPTX
  module OPC
    class PackageStreamer
      # Use ignore_part_exceptions to ignore exceptions while marshalling or streaming
      # a part. Useful to ignore missing files on S3.
      def initialize(package, ignore_part_exceptions=false)
        @package = package
        @ignore_part_exceptions = ignore_part_exceptions
      end

      # This method uses parts of zipline to implement streaming.
      #
      # This is how Zipline usually works:
      # * Rails calls response_body.each
      # * response_body is Zipline::ZipGenerator, it creates a Zipline::FakeStream with the
      #   block given to each
      # * The FakeStream is used as output for ZipLine::OutputStream (which inherits
      #   from Zip::OutputStream)
      # * ZipGenerator goes through files and calls write_file on the OutputStream for each file
      # * OutputStream writes the data as String to FakeStream via the << method
      # * FakeStream yields the String to the block given by Rails
      #
      # This method uses Zipline's FakeStream and OutputStream, but writes the entries
      # to the zip file on it's own, so we can write data for a file in multiple steps
      # and don't have to rely on Zipline's crazy logic (e.g. on what can be streamed and
      # what not - only IO and StringIO). Basically, this replaces ZipGenerator.
      def each(&block)
        fake_stream = Zipline::FakeStream.new(&block)
        Zipline::OutputStream.open(fake_stream) do |zip|
          @package.parts.each do |name, part|
            begin
              if part.respond_to?(:stream) && part.respond_to?(:size)
                zip.put_next_entry name, part.size
                part.stream zip
              else
                data = @package.marshal_part(name)
                zip.put_next_entry name, data.size
                zip << data
              end
            rescue => e
              raise e unless @ignore_part_exceptions
            end
          end
        end
      end
    end
  end
end
