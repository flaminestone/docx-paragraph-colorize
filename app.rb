require 'rubygems'
require 'chunky_png'
require 'image_size'
require 'rmagick'
filepath = 'image.bmp'
body = ''
image_size = ImageSize.new(File.open(filepath).read).size
wight = image_size.first
heading = image_size.last
pixels = Magick::ImageList.new(filepath).get_pixels(0, 0, wight, heading).each_slice(wight).to_a
font_size = 4
symbol = 'O'
symbols_line_size = 814
spacingline = 17

open('main.js', 'w') do |f|
  f << "builder.CreateFile(\"docx\");
var oDocument = Api.GetDocument();
var oParagraph, oRun;
oParagraph = oDocument.GetElement(0);
var oSection = oDocument.GetFinalSection();
oSection.SetPageMargins(0, 0, 0, 0);
oParagraph.SetSpacingLine(#{spacingline}, \"exact\");
var arr = ["
end


pixels.each_with_index do |current_column, pixel_number|
  symbols_line_size.times do |i|
   if i >= image_size.first
     body += '[250, 255, 255],'
   else
     color = current_column[i].to_color(Magick::AllCompliance, false, 8, true).match(/#(..)(..)(..)/)
     r = color[1].hex
     g = color[2].hex
     b = color[3].hex
     body += "[#{r}, #{g}, #{b}],"
   end
   puts(pixel_number)
 end
end
body = body[0...-1]
body += "];
var font_size = #{font_size};
var symbol = '#{symbol}';


arr.forEach(function(item, i, arr) {
    oRun = Api.CreateRun();
    oRun.SetFontSize(font_size);
    oRun.SetFontFamily(\"OpenSymbol\");
    oRun.SetColor(item[0], item[1], item[2]);
    oRun.AddText(symbol);
    oParagraph.AddElement(oRun)
});

"
open('main.js', 'a') { |f| f << body[0...-1] }
open('main.js', 'a') do |f|
  f << "builder.SaveFile(\"docx\", \"Result.docx\");
builder.CloseFile();"
end



`documentbuilder main.js`
