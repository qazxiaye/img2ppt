#!/usr/bin/ruby
# Copyright @ Ye XIA <qazxiaye@126.com>

require 'powerpoint'

if ARGV.length > 1
	puts "Usage:"
	puts "\033\[0;32m  ./img2ppt.rb\033\[0m : generate ppt using all the images in current folder as backgrounds"
	puts "  or"
	puts "\033\[0;32m  ./img2ppt.rb xxx.zip\033\[0m : unzip the zip file, and generate ppt based on unziped images in current foler"
	exit
end

dir = "#{Dir.pwd}"

if ARGV.length == 1
    time = Time.new
    time = "#{time.year}-#{time.month}-#{time.day}-#{time.hour}"
    dir = "#{dir}/#{time}"
	puts `unzip -o #{ARGV[0]} -d #{time}`
end

@deck = Powerpoint::Presentation.new
coords = {x: 0, y: 0, cx: 9147791, cy: 6870610}

isImg = false
for img in Dir["#{dir}/*.jpg"].sort
    @deck.add_pictorial_slide '', img, coords
	isImg = true
end
for img in Dir["#{dir}/*.png"].sort
    @deck.add_pictorial_slide '', img, coords
	isImg = true
end
for img in Dir["#{dir}/*.gif"].sort
    @deck.add_pictorial_slide '', img, coords
	isImg = true
end

if isImg
	@deck.save('test.pptx')
else
	puts "\033\[0;33mNo image found!\033\[0m"
end

if ARGV.length == 1
    puts `rm #{dir} -r`
end

