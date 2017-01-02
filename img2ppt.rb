# Copyright @ Ye XIA <qazxiaye@126.com>

require 'powerpoint'

@deck = Powerpoint::Presentation.new

coords = {x: 0, y: 0, cx: 9147791, cy: 6870610}

for img in Dir["#{Dir.pwd}/*.jpg"]
    @deck.add_pictorial_slide '', img, coords
end

@deck.save('test.pptx')

