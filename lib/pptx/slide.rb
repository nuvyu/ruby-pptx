module PPTX
  class Slide < OPC::Part
    def initialize(package, part_name)
      super(package, part_name, 'ppt/slides/slide1.xml')
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    end
  end
end
