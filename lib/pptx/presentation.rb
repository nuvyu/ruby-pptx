module PPTX
  class Presentation < OPC::Part
    def initialize(package, part_name)
      super(package, part_name)
    end

    def add_slide(slide)
      # TODO remove me - remove existing slide from template
      # slide_list_xml.xpath('./p:sldId').each do |slide|
      #   slide_list.children.delete(slide)
      # end
      ref_id = relationships.add(relative_part_name(slide.part_name), RELTYPE_SLIDE)

      # add slide to sldIdList in presentation
      slide_id = Nokogiri::XML::Node.new('p:sldId', doc)
      slide_id['id'] = next_slide_id
      slide_id['r:id'] = ref_id
      slide_list_xml.add_child(slide_id)
    end

    def next_slide_id
      ids = slide_list_xml.xpath('./p:sldId').map { |sid| sid['id'].to_i }
      return ids.max + 1
    end

    def slide_list_xml
      @slide_list_xml ||= doc.xpath('/p:presentation/p:sldIdLst', p: PRESENTATION_NS).first
    end
  end
end
