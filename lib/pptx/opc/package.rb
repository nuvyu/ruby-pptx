require 'nokogiri'
require 'zip'

module PPTX
  module OPC
    # Package allows the following:
    # * Create new package based on an existing PPTX file
    #   (extracted ZIP for better version control)
    # * Replace certain files with generated parts
    class Package
      # Create a new package using the directory strucuture base
      def initialize(base = nil)
        base ||= base || File.join(File.dirname(__FILE__), '..', 'base.pptx')
        @base = File.absolute_path(base)
        @parts = Hash[default_part_names.map {|fn| [fn, nil]}]
      end

      def base
        @base
      end

      def content_types
        @content_types ||= ContentTypes.new(self, '[Content_Types].xml')
      end

      def has_part?(name)
        @parts.has_key?(name)
      end

      def part(name)
        @parts[name]
      end

      def template_part(name)
        file = File.join(@base, name)
        IO.read(file) if File.exist? file
      end

      def marshal_part(name)
        part = @parts[name]
        if part
          if part.is_a? BasePart
            part.marshal
          elsif part.instance_of? String
            # TODO wrap binary parts, probably as references when implementing streaming
            part
          else
            raise "marshal_part: Unknown part type: '#{p.class}' for part '#{name}'"
          end
        else
          template_part name
        end
      end

      # Used e.g. for creating slides with proper numbering (i.e. slide1.xml, slide2.xml)
      def next_part_with_prefix(prefix, suffix)
        num = @parts.select { |n, _| n.start_with?(prefix) && n.end_with?(suffix) }
                   .map { |n, _| n[prefix.length..-1][0..-(suffix.size + 1)].to_i }
                   .max
        num = (num || 0) + 1
        "#{prefix}#{num}#{suffix}"
      end

      def set_part(name, obj, content_type = nil)
        @parts[name] = obj
        if content_type
          content_types.set_override(name, content_type)
        end
      end

      def presentation
        @presentation ||= PPTX::Presentation.new(self, 'ppt/presentation.xml')
      end

      class ZipStreamer
        def initialize(pkg)
          @pkg = pkg
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
            @pkg.instance_variable_get(:@parts).each do |name, _|
              part = @pkg.part(name)
              if part.respond_to?(:stream) && part.respond_to?(:size)
                zip.put_next_entry name, part.size
                part.stream zip
              else
                data = @pkg.marshal_part(name)
                zip.put_next_entry name, data.size
                zip << data
              end
            end
          end
        end
      end

      def stream_zip
        return ZipStreamer.new(self)
      end

      def to_zip
        buffer = ::StringIO.new('')

        Zip::OutputStream.write_buffer(buffer) do |out|
          @parts.each do |name, _|
            out.put_next_entry name
            out.write marshal_part(name)
          end
        end

        buffer.string
      end

      private
      def default_part_names
        @default_part_names ||= Dir.glob(File.join(@base, '**/*'), File::FNM_DOTMATCH)
                                   .reject {|fn| File.directory?(fn) }
                                   .reject {|fn| fn.end_with?('.DS_Store')}
                                   .map    {|fn| fn[(@base.length + 1)..-1]}
      end

    end
  end
end
