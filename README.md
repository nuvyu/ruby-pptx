# ruby-pptx

A ruby Gem for writing PPTX (Powerpoint) files. Right now, it's functionality is limited, it only allows adding textboxes with formatting and pictures to slides.

pptx allows streaming ZIPs while they're being created and allows streaming files and S3 objects into them. This allows you to create PPTX files without consuming large amounts of memory, either in RAM or on disk.

## Set up

Gemfile:

```ruby
gem 'pptx', git: 'https://github.com/nuvyu/ruby-pptx'
```

If you want to stream ZIP files, you'll need a yet unreleased version of [zipline](https://github.com/fringd/zipline). zipline isn't included as dependency for the Gem as it depends on Rails, which in turn adds a lot of dependencies and pptx is useful without Rails.

```ruby
# Used by pptx to stream zipfiles as a response to users
# Use a yet unreleased commit with a fix.
gem 'zipline', git: 'https://github.com/fringd/zipline',
               ref: 'd1e6e66e1c3e6945a3c09e90098f2f2f29e6c01e'
```

## Usage

**Note** As ruby-pptx' 0.x.x version indicates, its API isn't fixed and might change without announcement.

Here's an example that creates a presentation with one slide, two textboxes, and a picture. It writes the presentation to a file.

```ruby

pkg = PPTX::OPC::Package.new
slide = PPTX::Slide.new(pkg)

slide.add_textbox PPTX::cm(2, 1, 22, 2), 'Title :)', sz: 45*PPTX::POINT
slide.add_textbox PPTX::cm(14, 6, 10, 10), "Text with \nmultiple\nlines"

image = PPTX::OPC::FilePart.new(pkg, 'spec/fixtures/files/test_photo.jpg')
slide.add_picture PPTX::cm(2, 5, 10, 0.9), 'photo.jpg', image

pkg.presentation.add_slide(slide)

File.open('presentation.pptx', 'wb') {|f| f.write(pkg.to_zip) }
```

Things that might help you:

* Bounding boxes are given as an Array with four elements in English Metric Units (an OpenXML unit, defined e.g. as 1/360.000 centimeter). The format is `[left, top, width, height]`
* `PPTX::cm(2, 1, 22, 2)` allows you to convert multiple values from centimeter to EMUs.
* `Slide#add_picture` adds the picture with the given dimensions, so make sure to use the correct aspect ratio.
* `Slide#add_textbox` takes as third parameter a Hash of formatting options. The formatting values a are raw OOXML Text Run Properties (rPr) element attributes as defined in section 21.1.2.3.9 of part 1 of the spec linked below . Examples:
    * sz: font size (use PPTX::POINT)
    * b: 1 for bold
    * i: 1 for italic
* For development, the [Office OpenXML File Format specifications](http://www.ecma-international.org/publications/standards/Ecma-376.htm) might be useful
    - The 4th edition was used for developing this. See [part 1](http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fourth%20Edition,%20Part%201%20-%20Fundamentals%20And%20Markup%20Language%20Reference.zip) (warning: 5k pages PDF in a ZIP) for the document format and [part 4](http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fourth%20Edition,%20Part%202%20-%20Open%20Packaging%20Conventions.zip) for the package format.

