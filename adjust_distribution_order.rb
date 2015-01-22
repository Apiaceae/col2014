#encoding: utf-8

require 'nokogiri'
require 'open-uri'
require_relative "wrapper_dis_line.rb"


# user provide data file
from_file, to_file = ARGV


prompt = '>'

puts "please provide your input file name:", prompt
from_file = $stdin.gets.chomp

puts "please provide your output file name:", prompt
to_file = $stdin.gets.chomp




# read data into nokogir object
doc = Nokogiri::HTML(File.read(from_file))

seperator = "<br\>"

# read file content each line to array
line_arr = doc.to_s.split(seperator)

line_count = 0
sort_count = 0
modified_count = 0
line_prefix = "分布:"
dis_line = ''
# replace_string = ''
#result = ''

for line_string in line_arr
  line_text = line_string.to_s
  if line_text.include? line_prefix
    dis_line = line_text
    line_text = line_text.gsub(/^分布:/, '').gsub(/\s+/, '')

    if !line_text.include? ';' and !line_text.include? ','
      corrected_dis = dis_line
    end

    if line_text.include? ';' and line_text.include? ','
      corrected_dis = line_prefix + sort_province(array_province(line_text))+"; "+dis_arbord(line_text.gsub(/^分布:/, '')).to_s
    end

    if !line_text.include? ';' and line_text.include? ','
      if line_text.include? '原产' or line_text.include? '归化'
        corrected_dis = line_prefix+line_text
      else
        corrected_dis = line_prefix + sort_province(array_province(line_text))
      end
    end

    if line_text.include? ';' and !line_text.include? ','
      corrected_dis = line_prefix + sort_province(array_province(line_text))+"; "+dis_arbord(line_text.gsub(/^分布:/, '')).to_s
    end

    if line_text.gsub(/^分布:/, '').gsub(/\s+/, '') != corrected_dis.gsub(/^分布:/, '').gsub(/\s+/, '')
      modified_count = modified_count + 1
      # origin_dis = "分布:" + line_text.gsub(/;/, '; ')+" "
      origin_dis = "分布:" + line_text
      modified_dis = corrected_dis
      # replace_string = replace_string + "'" + origin_dis + "'" + " => " + "'" + modified_dis + "'" + ", "
      # org = "分布:" + line_text
      mod = modified_dis
      doc = doc.to_s.gsub(/#{dis_line}/, mod)
      # puts Nokogiri::HTML(doc).to_html
    end

    if need_sort(line_text)
      sort_count = sort_count + 1
      # puts line_count, line_text, corrected_dis
    end
    line_count = line_count + 1
    # puts "原始分布: #{origin_dis}"
    # puts "更正分布: #{modified_dis}"
    # ori = "/" + origin_dis + "/"
  end
end

# result = replace_string.chomp(', ')
# end

# puts "You have total #{line_count} row"
# puts "#{sort_count} records need sorted province order!"
puts "#{modified_count} records with modified distribution!"
# puts "#{result}"
docx = Nokogiri::HTML(doc).to_html
# puts docx
File.write(to_file, docx)
