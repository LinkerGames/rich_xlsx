# frozen_string_literal: true
require_relative "xml.rb"
require_relative "row.rb"

module RichXlsx
  class Worksheet
    def initialize(io, options = {})
      @io = io
      @rownum = 1
      @closed = false
      @options = options
      @has_header_row = options.delete(:has_header_row)
      @merge_cells = []
      write_header
    end

    def <<(row, common_cell_style = nil, custom_height = nil)
      options = @has_header_row && @rownum == 1 ? @options.merge(:is_header => true) : @options
      if common_cell_style.nil?
        @io << Row.new(row, @rownum, options).to_xml
      else
        if common_cell_style < 0
          raise "Invalid value for common_cell_style (#{common_cell_style}). Must be a non-negative integer obtained as a result of workbook.add_style({...})."
        end
        if custom_height.nil?
          @io << Row.new(row.map{|value| [value, common_cell_style]}, @rownum, options).to_xml
        else
          if custom_height < 0
            raise "Invalid custom_height value (#{custom_height})."
          end
          @io << Row.new(row.map{|value| [value, common_cell_style, custom_height]}, @rownum, options).to_xml
        end
      end
      @rownum += 1
    end
    alias_method :add_row, :<<

    def merge(top_left, bottom_right)
      @merge_cells << "#{top_left}:#{bottom_right}"
    end
    
    alias_method :merge_cells, :merge
    
    def close
      write_footer
      @closed = true
    end

    def closed?
      @closed
    end

    def id
      @options[:id]
    end

    def name
      @options[:name]
    end

    private

    def write_header
      @io << XML.header
      @io << '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'

      columns = Array(@options[:columns])
      unless columns.empty?
        @io << Columns.new(columns).to_xml
      end

      @io << '<sheetData>'
    end

    def write_footer
      @io << '</sheetData>'
      unless XML.blank?(@merge_cells)
        @io << '<mergeCells>'
        @merge_cells.each do |ref|
          @io << %(<mergeCell ref="#{XML.escape_attr(ref)}"/>)
        end
        @io << '</mergeCells>'
      end
      @io << '</worksheet>'
    end
  end
end
