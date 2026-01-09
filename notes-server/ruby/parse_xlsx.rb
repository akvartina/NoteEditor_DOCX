#!/usr/bin/env ruby
# frozen_string_literal: true

require "json"
require "zip"
require "nokogiri"

path = ARGV[0]
abort(JSON.generate({ ok: false, error: "Missing file path" })) unless path
abort(JSON.generate({ ok: false, error: "File not found: #{path}" })) unless File.exist?(path)

NS = {
  "wb" => "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "r"  => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "rel"=> "http://schemas.openxmlformats.org/package/2006/relationships"
}.freeze

def col_letters_to_index(letters)
  # "A"->0, "B"->1, ..., "Z"->25, "AA"->26
  sum = 0
  letters.each_byte do |b|
    sum = sum * 26 + (b - 64) # 'A' = 65
  end
  sum - 1
end

def parse_cell_ref(ref)
  # "C12" => [row_index=11, col_index=2]
  m = ref.match(/\A([A-Z]+)(\d+)\z/)
  return nil unless m
  col = col_letters_to_index(m[1])
  row = m[2].to_i - 1
  [row, col]
end

def read_entry(zip, name)
  e = zip.find_entry(name)
  return nil unless e
  e.get_input_stream.read
end

Zip::File.open(path) do |zip|
  workbook_xml = read_entry(zip, "xl/workbook.xml")
  abort(JSON.generate({ ok: false, error: "xl/workbook.xml missing" })) unless workbook_xml

  rels_xml = read_entry(zip, "xl/_rels/workbook.xml.rels")
  abort(JSON.generate({ ok: false, error: "xl/_rels/workbook.xml.rels missing" })) unless rels_xml

  shared_xml = read_entry(zip, "xl/sharedStrings.xml")
  shared_strings = []
  if shared_xml
    sdoc = Nokogiri::XML(shared_xml)
    # shared string items can have multiple <t> nodes (rich text); join them
    shared_strings = sdoc.xpath("//wb:sst/wb:si", NS).map do |si|
      si.xpath(".//wb:t", NS).map(&:text).join
    end
  end

  # Map rId -> Target worksheet path
  rdoc = Nokogiri::XML(rels_xml)
  rid_to_target = {}
  rdoc.xpath("//rel:Relationships/rel:Relationship", NS).each do |rel|
    rid = rel["Id"]
    target = rel["Target"]
    rid_to_target[rid] = target # e.g. "worksheets/sheet1.xml"
  end

  wdoc = Nokogiri::XML(workbook_xml)
  sheets_meta = []
  wdoc.xpath("//wb:workbook/wb:sheets/wb:sheet", NS).each do |sh|
    name = sh["name"]
    rid = sh.attribute_with_ns("id", NS["r"])&.value # r:id
    target = rid_to_target[rid]
    next unless target
    sheets_meta << { name: name, path: "xl/#{target}" }
  end

  sheets = sheets_meta.map do |meta|
    sheet_xml = read_entry(zip, meta[:path])
    next({ name: meta[:name], cells: [] }) unless sheet_xml

    sdoc = Nokogiri::XML(sheet_xml)

    # We'll build a sparse grid using a hash first
    grid = {} # row => { col => value }
    max_row = -1
    max_col = -1

    sdoc.xpath("//wb:worksheet/wb:sheetData/wb:row/wb:c", NS).each do |c|
      ref = c["r"] # e.g. "B2"
      rc = parse_cell_ref(ref)
      next unless rc
      r, col = rc

      cell_type = c["t"] # "s" (shared string), "b" (bool), "inlineStr", or nil (number)
      v = nil

      if cell_type == "s"
        idx = c.at_xpath("./wb:v", NS)&.text&.to_i
        v = idx ? shared_strings[idx] : ""
      elsif cell_type == "inlineStr"
        v = c.at_xpath("./wb:is/wb:t", NS)&.text.to_s
      elsif cell_type == "b"
        raw = c.at_xpath("./wb:v", NS)&.text.to_s
        v = (raw == "1") ? "TRUE" : "FALSE"
      else
        # number or empty
        raw = c.at_xpath("./wb:v", NS)&.text
        v = raw.nil? ? "" : raw
      end

      grid[r] ||= {}
      grid[r][col] = v

      max_row = r if r > max_row
      max_col = col if col > max_col
    end

    # Convert to 2D array (rows x cols)
    cells = []
    if max_row >= 0 && max_col >= 0
      (0..max_row).each do |r|
        row = Array.new(max_col + 1, "")
        if grid[r]
          grid[r].each { |col, val| row[col] = val }
        end
        cells << row
      end
    end

    { name: meta[:name], cells: cells }
  end.compact

  puts JSON.generate({ ok: true, sheets: sheets })
end