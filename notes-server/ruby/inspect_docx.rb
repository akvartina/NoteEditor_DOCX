#!/usr/bin/env ruby
# frozen_string_literal: true

require "ooxml_parser"

path = ARGV[0] or abort("Usage: ruby ruby/inspect_docx.rb file.docx")
doc = OoxmlParser::Parser.parse(path)

puts "DOC CLASS: #{doc.class}"

interesting = doc.methods.grep(/document|body|paragraph|table|header|footer|text|element/i).sort
puts "\nTop-level methods matching /document|body|paragraph|table|header|footer|text|element/:"
puts interesting.join(", ")

# Try common entry points and print what exists + counts
def try_path(label)
  obj = yield
  return if obj.nil?
  puts "\n== #{label} =="
  puts "class: #{obj.class}"
  if obj.respond_to?(:size)
    puts "size: #{obj.size}"
  elsif obj.respond_to?(:length)
    puts "length: #{obj.length}"
  end
  obj
rescue => e
  puts "\n== #{label} FAILED: #{e.class}: #{e.message}"
  nil
end

body = try_path("doc.document_body") { doc.document_body if doc.respond_to?(:document_body) }
try_path("body.paragraphs") { body.paragraphs if body && body.respond_to?(:paragraphs) }
try_path("doc.paragraphs") { doc.paragraphs if doc.respond_to?(:paragraphs) }
try_path("doc.elements") { doc.elements if doc.respond_to?(:elements) }
try_path("doc.document") { doc.document if doc.respond_to?(:document) }

# Find first bits of text anywhere by walking shallowly
def collect_text(obj, acc, depth)
  return if obj.nil? || depth <= 0 || acc.size >= 30

  if obj.is_a?(String)
    t = obj.strip
    acc << t if !t.empty?
    return
  end

  if obj.respond_to?(:text) && obj.text.is_a?(String)
    t = obj.text.strip
    acc << t if !t.empty?
  end

  if obj.is_a?(Array)
    obj.each { |x| collect_text(x, acc, depth - 1) }
    return
  end

  # Explore instance variables (works well for Ruby object graphs)
  obj.instance_variables.each do |iv|
    val = obj.instance_variable_get(iv)
    collect_text(val, acc, depth - 1)
    break if acc.size >= 30
  end
end

samples = []
collect_text(doc, samples, 6)

puts "\nText samples found by shallow walk (up to 30):"
puts samples.uniq.first(30).map { |s| "- #{s[0,120]}" }