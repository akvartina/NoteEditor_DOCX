#!/usr/bin/env ruby
# frozen_string_literal: true

require "json"
require "zip"
require "nokogiri"
require "fileutils"

EMU_PER_INCH = 914_400.0
def emu_to_in(v) = (v.to_f / EMU_PER_INCH)

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

def rels_map(zip, rels_path)
  xml = read_entry(zip, rels_path)
  return {} unless xml
  doc = Nokogiri::XML(xml)
  m = {}
  doc.xpath("//rel:Relationships/rel:Relationship", NS).each do |rel|
    m[rel["Id"]] = rel["Target"]
  end
  m
end

def slide_order(zip)
  pres_xml  = read_entry(zip, "ppt/presentation.xml")
  pres_rels = read_entry(zip, "ppt/_rels/presentation.xml.rels")
  raise "Missing presentation.xml" unless pres_xml
  raise "Missing presentation rels" unless pres_rels

  pres_doc = Nokogiri::XML(pres_xml)
  rels_doc = Nokogiri::XML(pres_rels)

  rid_to_target = {}
  rels_doc.xpath("//rel:Relationships/rel:Relationship", NS).each do |rel|
    rid_to_target[rel["Id"]] = rel["Target"] # slides/slide1.xml
  end

  pres_doc.xpath("//p:sldIdLst/p:sldId", NS).map do |sldId|
    rid = sldId["r:id"] || sldId.attributes.values.find { |a| a.name == "id" && a.namespace&.href == NS["r"] }&.value
    t = rid_to_target[rid]
    t ? "ppt/#{t}" : nil
  end.compact
end

def xfrm_in(pic_or_sp, ns)
  off = pic_or_sp.at_xpath(".//a:xfrm/a:off", ns)
  ext = pic_or_sp.at_xpath(".//a:xfrm/a:ext", ns)
  return [0,0,0,0] unless off && ext
  [emu_to_in(off["x"]), emu_to_in(off["y"]), emu_to_in(ext["cx"]), emu_to_in(ext["cy"])]
end

def extract_text(sp, ns)
  # join paragraphs with newline
  ps = sp.xpath(".//a:txBody/a:p", ns)
  lines = ps.map do |p|
    p.xpath(".//a:t", ns).map(&:text).join.strip
  end.reject(&:empty?)
  lines.join("\n")
end

# argv: pptx_path out_dir
pptx_path = ARGV[0]
out_dir   = ARGV[1]

abort(JSON.generate({ ok:false, error:"Usage: ruby parse_pptx_objects.rb file.pptx out_dir" })) unless pptx_path && out_dir
abort(JSON.generate({ ok:false, error:"File not found" })) unless File.exist?(pptx_path)

media_out = File.join(out_dir, "media")
FileUtils.mkdir_p(media_out)

Zip::File.open(pptx_path) do |zip|
  slides_paths = slide_order(zip)

  slides = slides_paths.each_with_index.map do |slide_path, idx|
    xml = read_entry(zip, slide_path)
    next nil unless xml
    doc = Nokogiri::XML(xml)

    # slide rels path: ppt/slides/_rels/slideN.xml.rels (N based on file name)
    m = /slide(\d+)\.xml/.match(slide_path)
    slide_num = m ? m[1] : nil
    
    rels_path = slide_num ? "ppt/slides/_rels/slide#{slide_num}.xml.rels" : nil
    rels = rels_path ? rels_map(zip, rels_path) : {}

    objects = []
    z = 0

    # Iterate spTree children in order -> z-order
    doc.xpath("//p:cSld/p:spTree/*", NS).each do |node|
      z += 1

      case node.name
      when "sp" # text shape
        text = extract_text(node, NS)
        next if text.strip.empty?
        x,y,w,h = xfrm_in(node, NS)
        objects << { id: "sp#{z}", type: "text", x: x, y: y, w: w, h: h, z: z, text: text }

      when "pic" # image
        rid = node.at_xpath(".//a:blip", NS)&.[]("r:embed")
        target = rid ? rels[rid] : nil
        next unless target

        # normalize ../media/imageX.png -> ppt/media/imageX.png
        media_rel = target.sub(/\A\.\.\//, "")
        media_path = "ppt/#{media_rel}"
        blob = read_entry(zip, media_path)
        next unless blob

        filename = File.basename(media_path)
        out_path = File.join(media_out, filename)
        File.binwrite(out_path, blob)

        x,y,w,h = xfrm_in(node, NS)
        objects << {
          id: "pic#{z}", type: "image",
          x: x, y: y, w: w, h: h, z: z,
          src: "/pptx-assets/#{File.basename(out_dir)}/media/#{filename}"
        }
      end
    end

    { index: idx + 1, objects: objects.sort_by { |o| o[:z] } }
  end.compact

  puts JSON.generate({ ok: true, slides: slides })
end