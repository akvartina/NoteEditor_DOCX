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
ns = { "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main" }

def esc(s)
  s.to_s.gsub("&", "&amp;").gsub("<", "&lt;").gsub(">", "&gt;")
end

def heading_level(p, ns)
  style = p.at_xpath("./w:pPr/w:pStyle/@w:val", ns)&.value
  return nil unless style
  # Common Word styles: Heading1, Heading2, ... or sometimes "Titre1" in localized templates
  if style =~ /Heading([1-6])/i
    $1.to_i
  else
    nil
  end
end

def run_to_html(r, ns)
  # Detect formatting
  bold = !r.xpath("./w:rPr/w:b", ns).empty?
  italic = !r.xpath("./w:rPr/w:i", ns).empty?
  underline = !r.xpath("./w:rPr/w:u", ns).empty?

  parts = []

  # Runs can contain multiple nodes: w:t (text), w:br (line break), w:tab (tab)
  r.children.each do |node|
    case node.name
    when "t"
      parts << esc(node.text)
    when "br"
      parts << "<br/>"
    when "tab"
      parts << "&nbsp;&nbsp;&nbsp;&nbsp;"
    end
  end

  s = parts.join
  return "" if s.strip.empty?

  s = "<u>#{s}</u>" if underline
  s = "<em>#{s}</em>" if italic
  s = "<strong>#{s}</strong>" if bold
  s
end

def paragraph_text_html(p, ns)
  runs = p.xpath(".//w:r", ns)
  html = runs.map { |r| run_to_html(r, ns) }.join
  html.strip
end

def is_list_paragraph?(p, ns)
  !p.xpath("./w:pPr/w:numPr", ns).empty?
end

out = []
in_ul = false

doc.xpath("//w:body/w:p", ns).each do |p|
  inner = paragraph_text_html(p, ns)
  next if inner.empty?

  if is_list_paragraph?(p, ns)
    out << "<ul>" unless in_ul
    in_ul = true
    out << "<li>#{inner}</li>"
    next
  end

  if in_ul
    out << "</ul>"
    in_ul = false
  end

  lvl = heading_level(p, ns)
  if lvl
    out << "<h#{lvl}>#{inner}</h#{lvl}>"
  else
    out << "<p>#{inner}</p>"
  end
end

out << "</ul>" if in_ul

html = out.join("\n")

puts JSON.generate({
  ok: true,
  html: html
})