require 'nokogiri'
require 'zip'

module PPTX
  module OPC
    # Package allows the following:
    # * Create new package based on an existing PPTX file
    #   (extracted ZIP for better version control)
    # * Replace certain files with generated versions
    class Package
      # Create a new package using the directory strucuture base
      def initialize(base = nil)
        base ||= base || File.join(File.dirname(__FILE__), '..', 'base.pptx')
        @base = File.absolute_path(base)
      end

      def base
        @base
      end

      def content_types
        @content_types = ContentTypes.new(self, '[Content_Types].xml')
      end

      def has_part?(name)
        parts.has_key?(name)
      end

      def base_part_xml(name)
        Nokogiri::XML(IO.read(File.join(@base, name)))
      end

      def part(name)
        parts[name] ||= IO.read(File.join(@base, name))
      end

      def marshal_part(name)
        unless has_part? name
          raise "marhsal_paprt: unknown part: #{name}"
        end

        part = parts[name]
        if part
          if part.is_a? BasePart
            part.marshal
          elsif part.instance_of? String
            part
          else
            raise "marshal_part: Unknown part type: '#{p.class}' for part '#{name}'"
          end
        else
          IO.read(File.join(@base, name))
        end
      end

      def set_part(name, obj, content_type)
        parts[name] = obj
        if content_type
          content_types.set_override(name, content_type)
        end
      end

      def part_xml(name)
        p = part(name)
        if p.is_a? BasePart
          p.doc
        elsif p.instance_of? String
          Nokogiri::XML(p)
        else
          raise "part_xml: Unknown part type: '#{p.class}' for part '#{name}'"
        end
      end

      def parts
        @parts ||= Hash[default_part_names.map {|fn| [fn, nil]}]
      end

      def default_part_names
        @default_part_names ||= Dir.glob(File.join(@base, '**/*'), File::FNM_DOTMATCH)
                                   .reject {|fn| File.directory?(fn) }
                                   .reject {|fn| fn.end_with?('.DS_Store')}
                                   .map    {|fn| fn[(@base.length + 1)..-1]}
      end

      def presentation
        @presentation ||= PPTX::Presentation.new(self, 'ppt/presentation.xml')
      end

      def to_zip
        buffer = ::StringIO.new('')

        Zip::OutputStream.write_buffer(buffer) do |out|
          parts.each do |name, content|
            out.put_next_entry name
            content ||= part(name)
            out.write marshal_part(name)
          end
        end

        buffer.string
      end
    end
  end
end
