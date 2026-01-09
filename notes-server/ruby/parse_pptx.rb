#!/usr/bin/env ruby
# frozen_string_literal: true

require "json"
require "zip"
require "nokogiri"

path = ARGV[0]
abort(JSON.generate({ ok: false, error: "Missing file path" })) unless path
abort(JSON.generate({ ok: false, error: "File not found: #{path}" })) unless File.exist?(path)

# Namespaces used in PPTX XML
NS = {
  "p" => "http://schemas.openxmlformats.org/presentationml/2006/main",
  "a" => "http://schemas.openxmlformats.org/drawingml/2006/main",
  "r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}.freeze

def read_entry(zip, name)
  e = zip.find_entry(name)
  return nil unless e
  e.get_input_stream.read
end

Zip::File.open(path) do |zip|
  # Find slide files like ppt/slides/slide1.xml, slide2.xml...
  slide_entries = zip.entries
                    .select { |e| e.name =~ %r{\Appt/slides/slide\d+\.xml\z} }
                    .sort_by { |e| e.name.scan(/\d+/).first.to_i }

  slides = slide_entries.map.with_index(1) do |entry, idx|
    xml = read_entry(zip, entry.name)
    next({ index: idx, text: "" }) unless xml

    doc = Nokogiri::XML(xml)

    # Extract all text runs <a:t> in reading order
    texts = doc.xpath("//a:t", NS).map(&:text).map(&:strip).reject(&:empty?)

    # Join into a simple slide text (one line per run/paragraph-ish)
    slide_text = texts.join("\n")

    { index: idx, text: slide_text }
  end.compact

  puts JSON.generate({ ok: true, slides: slides })
end