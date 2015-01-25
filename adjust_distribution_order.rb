#!/usr/bin/env ruby
# encoding: utf-8


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


# count how many distribution row need change
rows_need_change = 0
distribution = ''

# read data into nokogir object by html file input
# example_conifer.html
content = Nokogiri::HTML(File.read(from_file))


# read file content each line to array
arrays = content.to_s.split("<br\>")

arrays.each do |row|
  if row.include? '分布:'
    original_distribution = row.to_s
    row = row.to_s.gsub(/^分布:/, '').gsub(/\s+/, '')

    #sort with china and abroad distribution
    if row.include? ';' and row.include? ','
      distribution = '分布:' + sort_province(array_province(row))+"; "+dis_arbord(row).to_s
    end

    # sort for only china distribution
    if !row.include? ';' and row.include? ','
      if row.include? '原产' or row.include? '归化'
        distribution = '分布:' + row
      else
        distribution = '分布:' + sort_province(array_province(row))
      end
    end

    # sort for only one china distribution with abroad distribution
    if row.include? ';' and !row.include? ','
      distribution = '分布:' + sort_province(array_province(row))+"; "+dis_arbord(row).to_s
    end

    # only one distribution record don't need sort
    if !row.include? ';' and !row.include? ','
      distribution = '分布:' + row
    end

    # check if modified distribution sting is equal to original one
    if row.gsub(/\s+/, '') != distribution.gsub(/^分布:/, '').gsub(/\s+/, '')
      rows_need_change += 1
      row = '分布:' + row
      content = content.to_s.gsub(/#{original_distribution}/, distribution)
    end
  end
end

puts "#{rows_need_change} records with modified distribution!"

# puts "#{result}"
modified_content = Nokogiri::HTML(content).to_html

# puts docx
File.write(to_file, modified_content)
