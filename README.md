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

