#!/usr/bin/env ruby
# frozen_string_literal: true

require "json"
require "zip"
require "nokogiri"

path = ARGV[0]
abort(JSON.generate({ ok: false, error: "Missing file path" })) unless path
abort(JSON.generate({ ok: false, error: "File not found: #{path}" })) unless File.exist?(path)

xml = nil

Zip::File.open(path) do |zip|
  entry = zip.find_entry("word/document.xml")
  abort(JSON.generate({ ok: false, error: "word/document.xml not found in docx" })) unless entry
  xml = entry.get_input_stream.read
end

doc = Nokogiri::XML(xml)

# WordprocessingML namespaces
ns = { "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main" }

paragraphs = doc.xpath("//w:body/w:p", ns).map do |p|
  # Join all text nodes within the paragraph (including multiple runs)
  text = p.xpath(".//w:t", ns).map(&:text).join
  text.strip
end.reject(&:empty?)

# Basic HTML escaping
def esc(s)
  s
    .gsub("&", "&amp;")
    .gsub("<", "&lt;")
    .gsub(">", "&gt;")
end

html = paragraphs.map { |p| "<p>#{esc(p)}</p>" }.join("\n")

puts JSON.generate({
  ok: true,
  html: html,
  paragraph_count: paragraphs.length
  # test: paragraphs: paragraphs
})

