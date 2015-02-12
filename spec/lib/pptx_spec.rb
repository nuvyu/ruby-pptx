require 'spec_helper'

require 'pathname'
require 'pptx'
require 'stringio'
require 'zip'

describe 'PPTX' do
  describe 'with two textboxes and an image' do
    before do
      pkg = PPTX::OPC::Package.new
      slide = PPTX::Slide.new(pkg)

      slide.add_textbox([14*PPTX::CM, 6*PPTX::CM, 10*PPTX::CM, 10*PPTX::CM],
                        "Text box text\n:)\nwith<stuff>&to<be>escaped\nYay!")
      slide.add_textbox([2*PPTX::CM, 1*PPTX::CM, 22*PPTX::CM, 3*PPTX::CM],
                        'Title :)', sz: 45*PPTX::POINT)

      File.open('spec/fixtures/files/test_photo.jpg', 'r') do |image|
        slide.add_picture([2 * PPTX::CM, 5*PPTX::CM, 10*PPTX::CM, 10*PPTX::CM],
                        'photo.jpg', image)
      end

      pkg.presentation.add_slide(slide)

      slide2 = PPTX::Slide.new(pkg)
      pkg.presentation.add_slide(slide2)

      @zip_data = pkg.to_zip
    end


    # The following suddenly fails for unknown reasons.
    # Error also described here:
    # * https://github.com/rubyzip/rubyzip/issues/177
    # * https://github.com/rubyzip/rubyzip/issues/199
    # let(:zip) do
    #   io = StringIO.new(@zip_data)
    #   Zip::File.new(io, true, true).tap {|z| z.read_from_stream(io) }
    # end
    # let(:slide) { Nokogiri::XML(zip.read('ppt/slides/slide1.xml')) }
    # let(:slide_refs) { Nokogiri::XML(zip.read('ppt/slides/_rels/slide1.xml.rels')) }

    # This works around the problems:
    let(:zip_files) do
      zip_files = {}
      input_stream = Zip::InputStream.open(StringIO.new(@zip_data))
      while entry = input_stream.get_next_entry
        zip_files[entry.name] = entry.get_input_stream.read()
      end
      zip_files
    end

    it 'creates a zip file > 20k' do
      # Usually is about 50k, so make sure there's some stuff in there
      expect(@zip_data.size).to be > 20000
    end

    context 'slide 1' do
      let(:slide) { Nokogiri::XML(zip_files['ppt/slides/slide1.xml']) }

      let(:slide_refs) { Nokogiri::XML(zip_files['ppt/slides/_rels/slide1.xml.rels']) }

      let(:shape_tree) do
        slide.xpath('/p:sld/p:cSld/p:spTree', a:PPTX::DRAWING_NS, p:PPTX::Presentation::NS)
             .first
      end

      it 'contains two textboxes with the correct text' do
        textboxes = shape_tree.xpath('./p:sp/p:txBody')

        expect(textboxes.size).to eql(2)

        # Content textbox (first one)
        lines = textboxes[0].xpath('./a:p/a:r/a:t')
        expect(lines[0].content).to eql('Text box text')
        expect(lines[1].content).to eql(':)')
        expect(lines[2].content).to eql('with<stuff>&to<be>escaped')
        expect(lines[3].content).to eql('Yay!')

        # Title textbox (second one)
        expect(textboxes[1].xpath('./a:p/a:r/a:t').first.content).to eql 'Title :)'
      end

      it 'contains a reference to an image and the image itself' do
        pictures = shape_tree.xpath('./p:pic/p:blipFill/a:blip')

        expect(pictures.size).to eql(1)

        ref_id = pictures.first['r:embed']
        expect(ref_id).to start_with('rId')

        rels = slide_refs.xpath("/r:Relationships/r:Relationship[@Id='#{ref_id}']",
                                r: PPTX::OPC::Relationships::NS)
        expect(rels.size).to eq(1)

        image_part_ref = rels.first['Target']
        expect(image_part_ref).to start_with('../media/image-')
        expect(image_part_ref).to end_with('.jpg')

        image_part = Pathname.new('ppt/slides/').join(image_part_ref).cleanpath.to_s
        expect(zip_files[image_part].size). to be > 0
      end
    end

    it 'sets content type overrides for the slides' do
      doc = Nokogiri::XML(zip_files['[Content_Types].xml'])

      type_list = doc.xpath('/xmlns:Types').first

      ct = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

      override1 = type_list.xpath('xmlns:Override[@PartName="/ppt/slides/slide1.xml"]').first
      expect(override1['ContentType']).to eq ct

      override2 = type_list.xpath('xmlns:Override[@PartName="/ppt/slides/slide2.xml"]').first
      expect(override2['ContentType']).to eq ct
    end
  end
end
