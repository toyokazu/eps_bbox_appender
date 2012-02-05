#!/usr/bin/env ruby

require 'timeout'
require 'tempfile'

if ARGV.size < 1
  puts "eps_bbox_appender.rb:"
  puts "  reads a *.prn file created using Microsoft Office."
  puts "  then outputs EPS file with Bounding Box."
  puts "  When you outputs *.prn, use PS Printer Driver with"
  puts "  PostScript option as EPS."
  puts ""
  puts "usage: eps_bbox_appender.rb filename.prn [filename.prn ...]"
  exit 1
end

def get_bbox(file)
  begin
    pid = nil
    com = nil
    bbox = ""
    timeout(5) {
      com = IO.popen("gs -sDEVICE=bbox #{file} 2>&1", "r+")
      pid = com.pid
      com.puts "quit"
      while line = com.gets
        bbox << line if line =~ /^%%/
      end
    }
    return bbox
  rescue Timeout::Error => err
    Process.kill('SIGKILL', pid)
    com.close unless com.nil?
    return false
  end
end

ARGV.each do |file|
  prn_file = file
  eps_file = file.gsub(".prn", ".eps")
  tmp_file = Tempfile::new(file)
  open(prn_file) do |prn|
    while line = prn.gets
      next if line =~ /^#TPOGPS/
      tmp_file.print(line)
    end
  end
  tmp_file.close

  bbox = get_bbox(tmp_file.path)
  tmp_file.open
  if bbox
    open(eps_file, "w+b") do |eps|
      while line = tmp_file.gets
        if line =~ /^%\!PS-Adobe-3.0\sEPSF-3.0/
          eps.print(line)
          eps.print(bbox)
        else
          eps.print(line)
        end
      end
    end
  else
    $stderr.puts "can not extract Bounding Box of the file #{file}"
  end
  tmp_file.close
end
