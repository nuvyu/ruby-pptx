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
#
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

      def has_part?(name)
        parts.has_key?(name)
      end

      def part(name)
        parts[name] ||= IO.read(File.join(@base, name))
      end

      def part_xml(name)
        Nokogiri::XML(part(name))
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

      def save
        parts['[Content_Types].xml'] = @content_type_doc.to_s
      end

      def set_content_type(part_name, value)
        @content_type_doc ||= part_xml('[Content_Types].xml')
        type_list = @content_type_doc.xpath('c:Types', c: CONTENTTYPE_NS).first
        override = type_list.xpath("./xmlns:Override[@PartName='#{part_name}']").first
        unless override
          override = Nokogiri::XML::Node.new('Override', @content_type_doc)
          override['PartName'] = '/' + part_name
          type_list.add_child(override)
        end
        override['ContentType'] = value
      end

      def to_zip
        buffer = ::StringIO.new('')

        Zip::OutputStream.write_buffer(buffer) do |out|
          parts.each do |name, content|
            out.put_next_entry name
            content ||= part(name)
            out.write content
          end
        end

        buffer.string
      end
    end

    class Part
      def initialize(package, part_name)
        @package = package
        @part_name = part_name
      end

      def content_type
        nil
      end

      def doc
        @doc ||= @package.part_xml(@part_name)
      end

      def part_name
        @part_name
      end

      def relative_part_name(name)
        name[File.dirname(part_name).size + 1..-1]
      end

      def relationship_doc
        @relationship_doc ||= @package.part_xml(relationship_file)
      end

      def relationship_file
        File.join(File.dirname(@part_name), '_rels', File.basename(@part_name) + '.rels')
      end

      def save
        @package.parts[@part_name] = doc.to_s
        @package.parts[relationship_file] = relationship_doc.to_s
        if content_type
          @package.set_content_type(part_name, content_type)
        end
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
      ref_id = 'rId8'
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
      relationship_list = relationship_doc.xpath('r:Relationships', r: OPC::RELATIONSHIP_NS).first

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
      @package = package
      @part_name = part_name
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    end

    def doc
      @doc ||= default_doc @part_name, 'ppt/slides/slide1.xml'
    end

    def relationship_doc
      @relationship_doc ||= default_doc relationship_file, 'ppt/slides/_rels/slide1.xml.rels'
    end

    # Simplification for not wanting to fully generate all XML documents, but rather
    # load a template and modify it.
    def default_doc(part_name, default_file)
      if @package.has_part?(part_name)
        super
      else
        Nokogiri::XML(IO.read(File.join(@package.base, default_file)))
      end
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
  slide.save

  presentation = PPTX::Presentation.new(pkg, 'ppt/presentation.xml')
  presentation.add_slide(slide)
  presentation.save
  pkg.save

  pp pkg.parts

  File.open('tmp/generated.pptx', 'wb') {|f| f.write(pkg.to_zip) }
end

main
