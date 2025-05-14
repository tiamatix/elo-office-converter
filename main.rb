#####
# MIT License
# 
# Copyright (c) 2025 Wolf
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#####

# Description:
# This script converts an ELO-Office archive export to a flat directory structure.
#
# 1st export ELO office via the export dialog, make sure documents are exported, too.
# 2nd copy the export to a subfolder "in" in the same folder where this script is exported.
# 3rd run this script two times, 2nd run is required to update the folder time stamps

require 'fileutils'
require 'date'
require 'open3'

INPUT_ROOT = File.join(Dir.pwd, "in")
OUTPUT_ROOT = File.join(Dir.pwd, "out")

# look-up table for translated folder names
LOOK_UP = {}
LOOK_UP[INPUT_ROOT] = OUTPUT_ROOT

def parse_ini(path)
  ini = {}
  current_section = nil

  File.foreach(path, encoding: "Windows-1252") do |line|
    line.strip!
    next if line.empty? || line.start_with?(";", "#")

    if line =~ /^\[(.+?)\]$/
      current_section = Regexp.last_match(1)
      ini[current_section] = {}
    elsif line =~ /^([^=]+?)=(.*)$/
      key = Regexp.last_match(1).strip
      value = Regexp.last_match(2).strip
      ini[current_section][key] = value
    end
  end

  ini
end

def sanitize_filename(input)
  # Decode from Windows-1252 and ensure UTF-8
  utf8 = input.force_encoding("Windows-1252").encode("UTF-8", invalid: :replace, undef: :replace, replace: "_")

  # Replace German umlauts and ß
  replacements = {
    "Ä" => "Ae", "ä" => "ae",
    "Ö" => "Oe", "ö" => "oe",
    "Ü" => "Ue", "ü" => "ue",
    "ß" => "ss"
  }
  replaced = utf8.gsub(/[ÄäÖöÜüß]/, replacements)

  # Replace invalid characters (not allowed in file name)
  safe = replaced.gsub(/[\/\\:\*\?"<>\|]/, "_")
  safe = safe.gsub(/[[:cntrl:]]/, "").strip
  safe[0, 255].strip
end

def process_esw_file(full_path)
  relative_path = full_path.sub(/^#{Regexp.escape(INPUT_ROOT)}\/?/, "")
  output_path = File.join(OUTPUT_ROOT, relative_path)
  ini = parse_ini(full_path)
  desc_file_name = sanitize_filename(ini['GENERAL']['SHORTDESC'])
  current_folder = File.dirname(full_path)
  file_name_no_extension = File.basename(full_path, File.extname(full_path))  
  in_sub_folder = File.join(current_folder, file_name_no_extension)
  file_date = Date.strptime(ini['GENERAL']['ABLDATE'], "%Y-%m-%d")
  file_time = Time.new(file_date.year, file_date.month, file_date.day)
  if Dir.exist?(in_sub_folder)    
    out_sub_folder = File.join(LOOK_UP[current_folder], desc_file_name)
    LOOK_UP[in_sub_folder] = out_sub_folder
    puts "Processing folder: #{in_sub_folder} => #{out_sub_folder}"
    FileUtils.mkdir_p(out_sub_folder) unless Dir.exist?(out_sub_folder)
    Open3.capture3("touch", "-t", file_time.strftime("%Y%m%d%H%M"), out_sub_folder)
  else
    file_extension = ini['GENERAL']['DOCEXT']
    if file_extension.empty?
      puts "Empty file: #{full_path}"
      return
    end
    in_file_name = File.join(current_folder, "#{file_name_no_extension}#{file_extension}")
    out_folder = LOOK_UP[current_folder]
    out_file_name = File.join(out_folder, "#{desc_file_name}#{file_extension.downcase}")
    if !File.exist?(in_file_name)
      puts "missing file: #{in_file_name}"
      exit 1
    end
    puts "Processing file: #{in_file_name} => #{out_file_name}"
    FileUtils.cp(in_file_name, out_file_name) unless File.exist?(out_file_name)
    Open3.capture3("touch", "-t", file_time.strftime("%Y%m%d%H%M"), out_file_name)
  end
end

all_files = Dir.glob("#{INPUT_ROOT}/**/*.ESW")
# traverse down the files so parent folders are translated first
all_files.sort_by! { |d| d.count(File::SEPARATOR) }

all_files.each do |file|
  full_path = File.expand_path(file)
  process_esw_file(full_path)
end

puts "done."