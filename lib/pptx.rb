require 'bundler'
Bundler.require(:default, 'development')

module PPTX
  module OPC
    # Package allows the following:
    # * Create new package based on an existing PPTX file
    #   (extracted ZIP for better version control)
    # * Replace certain files with generated versions
    class Package
      # Create a new package using the directory strucuture base
      def initialize(base = nil)
        base ||= base || File.join(File.dirname(__FILE__), 'pptx', 'base.pptx')
        @base = File.absolute_path(base)
      end

      def part(name)
        parts[name] ||= IO.read(File.join(@base, name))
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

      def to_zip
        buffer = ::StringIO.new('')
        Zip::OutputStream.write_buffer(buffer) do |out|
          parts.each do |name, content|
            content ||= part(name)
            out.put_next_entry name
            out.write content
          end
        end
        buffer.string
      end
    end

    class Part
    end

    class Relationship
    end
  end
end



def main
  pkg = PPTX::OPC::Package.new

  app = Nokogiri::XML(pkg.part 'docProps/app.xml')
  slides = app.xpath('/ep:Properties/ep:Slides',
    ep: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties').first
  slides.content = 1
  pkg.parts['docProps/app.xml'] = app.to_s

  pp pkg.parts

  File.open('tmp/generated.pptx', 'wb') {|f| f.write(pkg.to_zip) }
end

main
