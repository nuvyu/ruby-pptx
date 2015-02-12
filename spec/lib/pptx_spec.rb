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

      # Remove existing pictures
      tree = slide.shape_tree_xml
      tree.xpath('./p:pic').each do |pic|
        pic.unlink
      end

      slide.add_textbox([14*PPTX::CM, 6*PPTX::CM, 10*PPTX::CM, 10*PPTX::CM],
                        "Text box text\n:)\nwith<stuff>&to<be>escaped\nYay!")
      slide.add_textbox([2*PPTX::CM, 1*PPTX::CM, 22*PPTX::CM, 3*PPTX::CM],
                        'Title :)', sz: 45*PPTX::POINT)

      File.open('spec/fixtures/files/test_photo.jpg', 'r') do |image|
        slide.add_picture([2 * PPTX::CM, 5*PPTX::CM, 10*PPTX::CM, 10*PPTX::CM],
                        'photo.jpg', image)
      end

      pkg.presentation.add_slide(slide)

      @zip_data = pkg.to_zip
    end

    let(:zip) do
      io = StringIO.new(@zip_data)
      Zip::File.new(io, true, true).tap {|z| z.read_from_stream(io) }
    end

    let(:slide) { Nokogiri::XML(zip.read('ppt/slides/slide2.xml')) }

    let(:slide_refs) { Nokogiri::XML(zip.read('ppt/slides/_rels/slide2.xml.rels')) }

    let(:shape_tree) do
      slide.xpath('/p:sld/p:cSld/p:spTree', a:PPTX::DRAWING_NS, p:PPTX::Presentation::NS)
           .first
    end

    it 'creates a zip file > 20k' do
      # Usually is about 50k, so make sure there's some stuff in there
      expect(@zip_data.size).to be > 20000
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
      expect(zip.read(image_part).size). to be > 0
    end
  end
end
