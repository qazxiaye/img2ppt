#!/usr/bin/ruby
# Copyright @ Ye XIA <qazxiaye@126.com>

require 'powerpoint'

if ARGV.length > 1
	puts "Usage:"
	puts "\033\[0;32m  ./img2ppt\033\[0m : generate ppt using all the images in current folder as backgrounds"
	puts "  or"
	puts "\033\[0;32m  ./img2ppt xxx.zip\033\[0m : unzip the zip file, and generate ppt based on unziped images in current foler"
	exit
end

if ARGV.length == 1
	puts `unzip -o #{ARGV[0]}`
end

@deck = Powerpoint::Presentation.new
coords = {x: 0, y: 0, cx: 9147791, cy: 6870610}

isImg = false
for img in Dir["#{Dir.pwd}/*.jpg"]
    @deck.add_pictorial_slide '', img, coords
	isImg = true
end
for img in Dir["#{Dir.pwd}/*.png"]
    @deck.add_pictorial_slide '', img, coords
	isImg = true
end
for img in Dir["#{Dir.pwd}/*.gif"]
    @deck.add_pictorial_slide '', img, coords
	isImg = true
end

if isImg
	@deck.save('test.pptx')
else
	puts "\033\[0;33mNo image found!\033\[0m"
end

