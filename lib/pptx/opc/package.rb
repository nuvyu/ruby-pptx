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

      # Returns a Hash of name -> Part
      # Do not modify the hash directly, but use set_part.
      def parts
        @parts
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

      # Stream a ZIP file while it's being generated
      # The returned PackageStreamer instance has an each method that takes a block, which
      # is called multiple times with chunks of data.
      # When adding BinaryParts (e.g. S3ObjectPart or FilePart), these are also
      # streamed into the ZIP file, so you can e.g. stream objects directly from S3
      # to the client, encoded in the ZIP file.
      # Use this e.g. in a Rails controller as `self.response_body`.
      def stream_zip(ignore_part_exceptions=false)
        PackageStreamer.new(self, ignore_part_exceptions)
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
