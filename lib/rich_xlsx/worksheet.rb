# frozen_string_literal: true
require_relative "errors.rb"
require_relative "xml.rb"
require_relative "row.rb"
require_relative "comments.rb"

module RichXlsx
  class Worksheet
    CELL_REFERENCE_REGEXP = /\A([A-Z]{1,3})([1-9][0-9]*)\z/.freeze

    attr_reader :comments

    def initialize(io, options = {})
      @io = io
      @rownum = 1
      @closed = false
      @options = options
      @has_header_row = options.delete(:has_header_row)
      @merge_cells = []
      @frozen_rows = 0
      @frozen_columns = 0
      @header_written = false
      @comments = []
      @comment_refs = {}
    end

    def <<(row, common_cell_style = nil, custom_height = nil)
      ensure_header
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
    
    # Adds a comment (note) to a cell, e.g. worksheet.add_comment('B6', 'Some explanation').
    # The cell may be in a row that has already been written or one that will be written later.
    # author: shown by Excel; defaults to the :comment_author workbook option.
    def add_comment(ref, text, author: nil)
      if @closed
        raise Error, "Comments can't be added to a closed worksheet."
      end
      ref = ref.to_s
      x, y = self.class.cell_to_xy(ref)
      text = text.to_s
      if text.empty?
        raise Error, "The comment text for cell #{ref} is empty."
      end
      if @comment_refs.key?(ref)
        raise Error, "Cell #{ref} already has a comment."
      end
      @comment_refs[ref] = true
      @comments << {ref: ref, x: x, y: y, text: text, author: author}.freeze
      nil
    end

    def freeze_pane(frozen_columns, frozen_rows)
      if @header_written
        raise "worksheet.freeze_pane can be called only before the header is written which is before the first worksheet row is added."
      end
      @frozen_columns = [frozen_columns.to_i, 0].max
      @frozen_rows = [frozen_rows.to_i, 0].max
    end
    
    def close
      ensure_header
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

    def self.xy_to_cell(x, y)
      col_index = x
      row_index = y
      col_letters = +""
      while col_index > 0
        col_index, remainder = (col_index - 1).divmod(26)
        col_letters.prepend((65 + remainder).chr)
      end
      "#{col_letters}#{row_index}"
    end
    
    # Converts "B3" to [2, 3] (1-based column and row). The reverse of xy_to_cell.
    def self.cell_to_xy(ref)
      match = CELL_REFERENCE_REGEXP.match(ref.to_s)
      if match.nil?
        raise Error, "Invalid cell reference #{ref.inspect}."
      end
      column = match[1].each_char.reduce(0) { |acc, ch| (acc * 26) + (ch.ord - 64) }
      return column, match[2].to_i
    end

    private
    
    def ensure_header
      if @header_written
        return
      end
      @io << XML.header
      @io << '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'

      if (@frozen_columns > 0) || (@frozen_rows > 0)
        @io << %(<sheetViews><sheetView workbookViewId="0"><pane xSplit="#{@frozen_columns}" ySplit="#{@frozen_rows}" topLeftCell="#{self.class.xy_to_cell(@frozen_columns + 1, @frozen_rows + 1)}" state="frozen"/></sheetView></sheetViews>)
      end
      
      columns = Array(@options[:columns])
      unless columns.empty?
        @io << Columns.new(columns).to_xml
      end

      @io << '<sheetData>'
      @header_written = true
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
      unless @comments.empty?
        # legacyDrawing is one of the last elements of a worksheet - it must come after mergeCells.
        @io << %(<legacyDrawing r:id="#{Comments::VML_DRAWING_RELATIONSHIP_ID}"/>)
      end
      @io << '</worksheet>'
    end
  end
end
