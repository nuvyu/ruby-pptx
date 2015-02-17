require 'bundler'

# Bundler setup instead of require and doing manual requires significantly speeds up startup
# as not the whole project's dependencies are loaded.
Bundler.setup(:default, 'development')

require_relative '../pptx'


def main
  pkg = PPTX::OPC::Package.new
  slide = PPTX::Slide.new(pkg)

  slide.add_textbox([14*PPTX::CM, 6*PPTX::CM, 10*PPTX::CM, 10*PPTX::CM],
                    "Text box text\n:)\nwith<stuff>&to<be>escaped\nYay!")
  slide.add_textbox([2*PPTX::CM, 1*PPTX::CM, 22*PPTX::CM, 3*PPTX::CM],
                    'Title :)', sz: 45*PPTX::POINT)

  image = PPTX::OPC::FilePart.new(pkg, 'spec/fixtures/files/test_photo.jpg')
  slide.add_picture([2 * PPTX::CM, 5*PPTX::CM, 10*PPTX::CM, 10*PPTX::CM],
                  'photo.jpg', image)

  pkg.presentation.add_slide(slide)

  slide2 = PPTX::Slide.new(pkg)
  pkg.presentation.add_slide(slide2)

  File.open('tmp/generated.pptx', 'wb') {|f| f.write(pkg.to_zip) }
end

main
