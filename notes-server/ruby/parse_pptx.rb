#!/usr/bin/env ruby
# frozen_string_literal: true

require "json"
require "zip"
require "nokogiri"

path = ARGV[0]
abort(JSON.generate({ ok: false, error: "Missing file path" })) unless path
abort(JSON.generate({ ok: false, error: "File not found: #{path}" })) unless File.exist?(path)

NS = {
  "p" => "http://schemas.openxmlformats.org/presentationml/2006/main",
  "a" => "http://schemas.openxmlformats.org/drawingml/2006/main",
  "r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "rel" => "http://schemas.openxmlformats.org/package/2006/relationships"
}.freeze

def read_entry(zip, name)
  e = zip.find_entry(name)
  return nil unless e
  e.get_input_stream.read
end

def paragraph_text(p_node, ns)
  parts = []
  p_node.children.each do |ch|
    case ch.name
    when "r"
      t = ch.at_xpath(".//a:t", ns)&.text
      parts << t.to_s if t
    when "br"
      parts << "\n"
    end
  end
  parts.join.strip
end

def paragraph_level(p_node, ns)
  lvl = p_node.at_xpath("./a:pPr", ns)&.[]("lvl")
  (lvl ? lvl.to_i : 0)
end

def placeholder_type(sp_node, ns)
  sp_node.at_xpath("./p:nvSpPr/p:nvPr/p:ph", ns)&.[]("type")
end

def extract_textbody_blocks(sp_node, ns)
  tx = sp_node.at_xpath(".//a:txBody", ns)
  return [] unless tx

  tx.xpath("./a:p", ns).map do |p|
    {
      level: paragraph_level(p, ns),
      text: paragraph_text(p, ns)
    }
  end.reject { |x| x[:text].nil? || x[:text].empty? }
end

Zip::File.open(path) do |zip|
  pres_xml  = read_entry(zip, "ppt/presentation.xml")
  pres_rels = read_entry(zip, "ppt/_rels/presentation.xml.rels")
  abort(JSON.generate({ ok: false, error: "ppt/presentation.xml missing" })) unless pres_xml
  abort(JSON.generate({ ok: false, error: "ppt/_rels/presentation.xml.rels missing" })) unless pres_rels

  pres_doc = Nokogiri::XML(pres_xml)
  rels_doc = Nokogiri::XML(pres_rels)

  rid_to_target = {}
  rels_doc.xpath("//rel:Relationships/rel:Relationship", NS).each do |rel|
    rid_to_target[rel["Id"]] = rel["Target"] # e.g. "slides/slide1.xml"
  end

  slide_targets = pres_doc.xpath("//p:sldIdLst/p:sldId", NS).map do |sldId|
    rid = sldId.attribute_with_ns("id", NS["r"]) # often nil
    rid = sldId.attribute_with_ns("id", NS["r"])&.value
    # Nokogiri can be finicky with attribute_with_ns; fallback:
    rid ||= sldId.attribute("r:id")&.value
    rid ||= sldId.attributes.values.find { |a| a.name == "id" && a.namespace&.href == NS["r"] }&.value
    rid ||= sldId.attributes["id"]&.value # last resort
    target = rid_to_target[rid]
    target ? "ppt/#{target}" : nil
  end.compact

  slides = []
  slide_targets.each_with_index do |slide_path, idx|
    xml = read_entry(zip, slide_path)
    next unless xml

    doc = Nokogiri::XML(xml)

    title = nil
    blocks = []

    # Each text-containing shape is p:sp
    doc.xpath("//p:cSld/p:spTree/p:sp", NS).each do |sp|
      ph = placeholder_type(sp, NS) # "title", "ctrTitle", "body", etc.
      paras = extract_textbody_blocks(sp, NS)
      next if paras.empty?

      if ph == "title" || ph == "ctrTitle"
        # First non-empty paragraph becomes title
        title ||= paras.map { |x| x[:text] }.find { |t| !t.empty? }
        next
      end

      if ph == "body"
        # Treat as bullets (keep levels)
        items = paras.map { |x| { level: x[:level], text: x[:text] } }
        blocks << { type: "bullets", items: items } unless items.empty?
      else
        # Other textboxes: join paragraphs with newlines
        text = paras.map { |x| x[:text] }.join("\n")
        blocks << { type: "text", text: text } unless text.strip.empty?
      end
    end

    slides << {
      index: idx + 1,
      title: title.to_s,
      blocks: blocks
    }
  end

  puts JSON.generate({ ok: true, slides: slides })
end