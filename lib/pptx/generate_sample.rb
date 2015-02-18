require 'bundler'

# Bundler setup instead of require and doing manual requires significantly speeds up startup
# as not the whole project's dependencies are loaded.
Bundler.setup(:default, 'development')

require_relative '../pptx'


def main
  pkg = PPTX::OPC::Package.new
  slide = PPTX::Slide.new(pkg)

  title_dimensions = PPTX::cm(2, 1, 22, 2)
  date_dimensions = PPTX::cm(2, 2.8, 5, 2)
  text_dimensions = PPTX::cm(14, 6, 10, 10)
  image_dimensions = PPTX::cm(2, 5, 10, 10)

  slide.add_textbox title_dimensions, 'Title :)', sz: 45*PPTX::POINT
  slide.add_textbox date_dimensions, '18 Feb 2015'
  slide.add_textbox text_dimensions,
                    "Text box text with long text that should be broken " +
                    "down into multiple lines.\n:)\nwith<stuff>&to<be>escaped\nYay!"

  image = PPTX::OPC::FilePart.new(pkg, 'spec/fixtures/files/test_photo.jpg')
  slide.add_picture image_dimensions, 'photo.jpg', image

  pkg.presentation.add_slide(slide)

  slide2 = PPTX::Slide.new(pkg)
  slide2.add_textbox title_dimensions,
                     'Very long title that should be broken up into multiple' +
                     'lines and possibly clipped', sz: 45*PPTX::POINT
  slide2.add_textbox date_dimensions, '17 Feb 2015'
  slide2.add_textbox text_dimensions,
                     "Text box text with long text that should be broken " +
                     "up into multiple lines and should overflow overflow" +
                     " overflowoverflowoverflowoverflowoverflowoverflow" +
                     " overflow overflow overflow overflow overflow overflow" +
                     " overflow overflow overflow overflow overflow overflow" +
                     " overflow overflow overflow overflow overflow overflow" +
                     " overflow overflow overflow overflow overflow overflow" +
                     " overflow overflow overflow overflow overflow overflow" +
                     " overflow overflow overflow overflow overflow overflow" +
                     " overflow overflow overflow overflow overflow overflow"
  pkg.presentation.add_slide(slide2)

  File.open('tmp/generated.pptx', 'wb') {|f| f.write(pkg.to_zip) }
end

main
