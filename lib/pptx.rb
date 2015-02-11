require 'bundler'

# Bundler setup instead of require and doing manual requires significantly speeds up startup
# as not the whole project's dependencies are loaded.
Bundler.setup(:default, 'development')

require 'nokogiri'
require 'zip'
require 'pp'
require 'securerandom'

# Assumptions
# * XML namespace prefixes remain the same as in template
# * Main presentation is in ppt/presentation.xml
# * Base presentation is trusted. Nothing is done to prevent directory traversal.
# If you wanted to open arbitrary presentation, you'd have to make sure to no longer make these.
module PPTX
  module OPC
    RELATIONSHIP_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    CONTENTTYPE_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
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

    class Relationships < BasePart

    end

    class ContentTypes < BasePart
      def set_override(part_name, value)
        type_list = doc.xpath('c:Types', c: CONTENTTYPE_NS).first
        override = type_list.xpath("./xmlns:Override[@PartName='#{part_name}']").first
        unless override
          override = Nokogiri::XML::Node.new('Override', doc)
          override['PartName'] = '/' + part_name
          type_list.add_child(override)
        end
        override['ContentType'] = value
      end
    end
  end

  RELATIONSHIP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  PRESENTATION_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
  RELTYPE_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'

  class Presentation < OPC::Part
    def initialize(package, part_name)
      super(package, part_name)
    end

    def add_slide(slide)
      slide_list = slide_list_xml
      ref_id = "rId#{SecureRandom.hex(10)}"
      # slide_part_name = 'slides/slide1.xml'

      # TODO remove me - remove existing slide from template
      # slide_list.xpath('./p:sldId').each do |slide|
      #   slide_list.children.delete(slide)
      # end

      # add slide to sldIdList in presentation
      slide_id = Nokogiri::XML::Node.new('p:sldId', doc)
      slide_id['id'] = next_slide_id
      slide_id['r:id'] = ref_id
      slide_list.add_child(slide_id)

      # add slide to presentation relationships, but first clear existing slides
      relationship_list = relationships.doc.xpath('r:Relationships', r: OPC::RELATIONSHIP_NS).first

      # TODO remove me - remove existing slide from template
      # relationship_list.xpath("./xmlns:Relationship[@Type='#{RELTYPE_SLIDE}']").each do |r|
      #   relationship_list.children.delete(r)
      # end
      relationship = Nokogiri::XML::Node.new('Relationship', doc)
      relationship['Id'] = ref_id

      relationship['Target'] = relative_part_name(slide.part_name)
      relationship['Type'] = RELTYPE_SLIDE
      relationship_list.add_child(relationship)
    end

    def next_slide_id
      ids = slide_list_xml.xpath('./p:sldId').map { |sid| sid['id'].to_i }
      return ids.max + 1
    end

    def slide_list_xml
      doc.xpath('/p:presentation/p:sldIdLst', p: PRESENTATION_NS).first
    end
  end

  class Slide < OPC::Part
    def initialize(package, part_name)
      super(package, part_name, 'ppt/slides/slide1.xml')
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    end
  end
end



def main
  pkg = PPTX::OPC::Package.new

  # TODO move this to presentation
  # app = pkg.part_xml 'docProps/app.xml'
  # slides = app.xpath('/ep:Properties/ep:Slides',
  #   ep: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties').first
  # slides.content = 1
  # pkg.parts['docProps/app.xml'] = app.to_s

  slide = PPTX::Slide.new(pkg, 'ppt/slides/slide2.xml')

  presentation = pkg.presentation
  presentation.add_slide(slide)

  File.open('tmp/generated.pptx', 'wb') {|f| f.write(pkg.to_zip) }
end

main
