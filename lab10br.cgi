#!/usr/bin/env ruby
require 'cgi'

cgi = CGI.new
city = cgi['city'].capitalize
province = cgi['province'].capitalize unless cgi['province'].empty?
country = cgi['country'].capitalize
image_url = cgi['image_url']

puts "Content-type: text/html\n\n"
puts "<html>"
puts "<head><title>City Information</title></head>"
puts "<body style='text-align: center; background-color: lightyellow;'>"
puts "<h1 style='font-size: 3em; color: blue;'>#{city}, #{province}, #{country}</h1>" unless province.nil?
puts "<h1 style='font-size: 3em; color: blue;'>#{city}, #{country}</h1>" if province.nil?
puts "<img src='#{image_url}' style='width: 100%; height: auto;'>"
puts "</body>"
puts "</html>"
