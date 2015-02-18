require 'bundler'

# Bundler setup instead of require and doing manual requires significantly speeds up startup
# as not the whole project's dependencies are loaded.
Bundler.setup(:default, 'development')

require_relative '../pptx'


def main
  pkg = PPTX::OPC::Package.new
  slide = PPTX::Slide.new(pkg)

  title_dimensions = [2*PPTX::CM, 1*PPTX::CM, 22*PPTX::CM, 2*PPTX::CM]
  date_dimensions = [2*PPTX::CM, 2.8*PPTX::CM, 5*PPTX::CM, 2*PPTX::CM]
  text_dimensions = [14*PPTX::CM, 6*PPTX::CM, 10*PPTX::CM, 10*PPTX::CM]
  image_dimensions = [2 * PPTX::CM, 5*PPTX::CM, 10*PPTX::CM, 10*PPTX::CM]

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
