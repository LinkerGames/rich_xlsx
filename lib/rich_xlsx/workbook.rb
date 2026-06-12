# frozen_string_literal: true
require "set"
require_relative "errors.rb"
require_relative "xml.rb"
require_relative "shared_string_table.rb"
require_relative "worksheet.rb"
require_relative "zip_kit_writer.rb"

module RichXlsx
  class Workbook

    VALID_BORDER_KEYS = Set.new([
      :left,
      :right,
      :top,
      :bottom,
      :diagonal
    ]).freeze
    
    VALID_OUTLINE_BORDER_KEYS = [
      :left,
      :right,
      :top,
      :bottom
    ].freeze
    
    VALID_BORDER_STYLES = Set.new([
      'thin',
      'medium',
      'thick',
      'dashed',
      'dotted',
      'double',
      'hair',
      'mediumDashed',
      'dashDot',
      'mediumDashDot',
      'dashDotDot',
      'mediumDashDotDot',
      'slantDashDot',
    ]).freeze
    
    VALID_HORIZONTAL_ALIGNMENTS = Set.new([
      'general',
      'left',
      'center',
      'right',
      'fill',
      'justify',
      'centerContinuous',
      'distributed'
    ]).freeze
    
    VALID_VERTICAL_ALIGNMENTS = Set.new([
      'top',
      'center',
      'bottom',
      'justify',
      'distributed'
    ]).freeze
    
    class << self

      def open(output, options = {})
        workbook = new(output, options)
        if block_given?
          begin
            yield workbook
          ensure
            workbook.close
          end
        else
          workbook
        end
      end

    end

    def initialize(output, options = {})
      @writer = ZipKitWriter.with_output_to(output)
      @options = options
      @sst = SharedStringTable.new
      @worksheets = []
      @formats = [
        "yyyy\\-mm\\-dd",
        "yyyy\\-mm\\-dd hh:mm:ss"
      ]
      @fonts = [
        {
          size: @options.fetch(:font, {}).fetch(:size, 11),
          name: @options.fetch(:font, {}).fetch(:name, 'Calibri'),
        }.freeze,
        {
          size: @options.fetch(:font, {}).fetch(:size, 11),
          name: @options.fetch(:font, {}).fetch(:name, 'Calibri'),
          bold: true
        }.freeze
      ]
      @fills = [
        {
          patternType: 'none'
        }.freeze,
        {
          patternType: 'gray125'
        }.freeze
      ]
      @borders = [{}.freeze]
      @styles = [
        {}.freeze,
        {
          numFmtId: 164,
        }.freeze,
        {
          numFmtId: 165,
        }.freeze,
        {
          fontId: 1,
          h_align: 'center',
          v_align: 'center'
        }.freeze,
        {
          numFmtId: 164,
          fontId: 1,
          h_align: 'center',
          v_align: 'center'
        }.freeze,
        {
          numFmtId: 165,
          fontId: 1,
          h_align: 'center',
          v_align: 'center'
        }.freeze
      ]
      @worksheed_added = false
    end

    def add_style(hash)
      if @worksheed_added
        fail Error, "Calling add_style after calling write_worksheet is not allowed."
      end
      style = {}
      [
        [:format, @formats, :numFmtId],
        [:font, @fonts, :fontId],
        [:fill, @fills, :fillId],
        [:border, @borders, :borderId]
      ].each do |tuple|
        value_hash = hash[tuple[0]]
        arr = tuple[1]
        style_key = tuple[2]
        unless XML.blank?(value_hash)
          value_index = arr.index(value_hash)
          if value_index.nil?
            arr << value_hash
            value_index = (arr.size - 1)
          end
          if style_key == :numFmtId
            value_index += 164
          end
          style[style_key] = value_index
        end
      end
      h_align = hash[:h_align]
      unless XML.blank?(h_align)
        if VALID_HORIZONTAL_ALIGNMENTS.include?(h_align)
          style[:h_align] = h_align
        else
          raise "Invalid value for h_align #{h_align.inspect}"
        end
      end
      v_align = hash[:v_align]
      unless XML.blank?(v_align)
        if VALID_VERTICAL_ALIGNMENTS.include?(v_align)
          style[:v_align] = v_align
        else
          raise "Invalid value for v_align #{v_align.inspect}"
        end
      end
      word_wrap = hash[:word_wrap]
      if word_wrap == true
        style[:word_wrap] = true
      end
      style.freeze
      style_index = @styles.index(style)
      if style_index.nil?
        @styles << style
        style_index = (@styles.size - 1)
      end
      return style_index
    end
    
    def add_worksheet(*args, &block)
      if block_given?
        # This method used to be an alias for `write_worksheet`. This was never publicly documented,
        # but to avoid breaking this private API we keep the old behaviour when called with a block.
        Kernel.warn "#{caller.first[/.*:\d+:(?=in `)/]} warning: Calling #{self.class}#add_worksheet with a block is deprecated, use #write_worksheet instead."
        return write_worksheet(*args, &block)
      end

      unless @worksheets.all? { |ws| ws.closed? }
        fail Error, "Close the current worksheet before adding a new one"
      end

      build_worksheet(*args)
    end

    def write_worksheet(*args)
      worksheet = build_worksheet(*args)

      yield worksheet if block_given?
      worksheet.close

      nil
    end

    def close
      write_workbook
      write_styles
      write_sst unless @sst.empty?
      write_workbook_rels
      write_root_rels
      write_content_types
      @writer.close

      nil
    end

    private
    def build_worksheet(name = nil, options = {})
      if name.is_a? Hash and options.empty?
        options = name
        name = nil
      end

      use_sst = options.fetch(:use_shared_strings, @options[:use_shared_strings])
      auto_format = options.fetch(:auto_format, @options[:auto_format])
      columns = options.fetch(:columns, @options[:columns])
      has_header_row = options.fetch(:has_header_row, @options[:has_header_row])
      sst = use_sst ? @sst : nil

      sheet_id = @worksheets.size + 1
      name = name || options[:name] || "Sheet#{sheet_id}"

      @writer.add_file "xl/worksheets/sheet#{sheet_id}.xml"

      worksheet = Worksheet.new(@writer, :id => sheet_id, :name => name, :sst => sst, :auto_format => auto_format, :columns => columns, :has_header_row => has_header_row)
      @worksheets << worksheet
      @worksheed_added = true
      worksheet
    end

    def write_root_rels
      @writer.add_file "_rels/.rels"
      @writer << XML.header
      @writer << XML.strip(<<-XML)
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
      XML
    end

    def write_workbook
      rid = String.new("rId0")
      @writer.add_file "xl/workbook.xml"
      @writer << XML.header
      @writer << XML.strip(<<-XML)
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <workbookPr date1904="false"/>
          <sheets>
      XML
      @worksheets.each do |worksheet|
        @writer << %'<sheet name="#{XML.escape_attr worksheet.name}" sheetId="#{worksheet.id}" r:id="#{rid.next!}"/>'
      end
      @writer << XML.strip(<<-XML)
          </sheets>
        </workbook>
      XML
    end

    def write_styles
      @writer.add_file "xl/styles.xml"
      @writer << XML.header
      @writer << XML.strip(<<-XML)
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <numFmts count="#{@formats.size}">
      XML
      @formats.each_with_index do |fmt, index|
        @writer << %(<numFmt numFmtId="#{164 + index}" formatCode="#{XML.escape_attr(fmt)}"/>)
      end
      @writer << XML.strip(<<-XML)
          </numFmts>
          <fonts count="#{@fonts.size}">
      XML
      @fonts.each do |font|
        @writer << '<font>'
        @writer << '<b/>' if font[:bold]
        @writer << '<i/>' if font[:italic]
        @writer << '<u/>' if font[:underline]
        @writer << '<strike/>' if font[:strike]
        unless XML.blank?(font[:color])
          @writer << %(<color rgb="#{XML.escape_attr(font[:color])}"/>)
        end
        unless XML.blank?(font[:size])
          @writer << %(<sz val="#{XML.escape_attr(font[:size].to_s)}"/>)
        end
        unless XML.blank?(font[:name])
          @writer << %(<name val="#{XML.escape_attr(font[:name])}"/>)
        end
        @writer << '</font>'
      end
      @writer << XML.strip(<<-XML)
          </fonts>
          <fills count="#{@fills.size}">
      XML
      @fills.each do |fill|
        @writer << '<fill><patternFill'
        @writer << %( patternType="#{fill[:patternType]}")
        unless XML.blank?(fill[:fgColor]) && XML.blank?(fill[:bgColor])
          @writer << '>'
          unless XML.blank?(fill[:fgColor])
            @writer << %(<fgColor rgb="#{XML.escape_attr(fill[:fgColor])}"/>)
          end
          unless XML.blank?(fill[:bgColor])
            @writer << %(<bgColor rgb="#{XML.escape_attr(fill[:bgColor])}"/>)
          end
          @writer << '</patternFill></fill>'
        else
          @writer << '/></fill>'
        end
      end
      @writer << XML.strip(<<-XML)
          </fills>
          <borders count="#{@borders.size}">
      XML
      @borders.each do |border|
        if XML.blank?(border)
          @writer << '<border/>'
        else
          @writer << '<border'
          unless XML.blank?(border[:diagonal])
            unless XML.blank?(border[:diagonal][:diagonalUp])
              @writer << %( diagonalUp="#{border[:diagonal][:diagonalUp] ? '1' : '0'}")
            end
            unless XML.blank?(border[:diagonal][:diagonalDown])
              @writer << %( diagonalDown="#{border[:diagonal][:diagonalDown] ? '1' : '0'}")
            end
          end
          @writer << '>'
          VALID_OUTLINE_BORDER_KEYS.each do |key|
            unless XML.blank?(border[key])
              @writer << %(<#{key})
              unless XML.blank?(border[key][:style])
                @writer << %( style="#{border[key][:style]}")
              end
              @writer << '>'
              unless XML.blank?(border[key][:color])
                @writer << %(<color rgb="#{XML.escape_attr(border[key][:color])}"/>)
              end
              @writer << %(</#{key}>)
            end
          end
          unless XML.blank?(border[:diagonal]) || (XML.blank?(border[:diagonal][:diagonalUp]) && XML.blank?(border[:diagonal][:diagonalDown]))
            @writer << '<diagonal'
            unless XML.blank?(border[:diagonal][:style])
              @writer << %( style="#{border[:diagonal][:style]}")
            end
            @writer << '>'
            unless XML.blank?(border[:diagonal][:color])
              @writer << %(<color rgb="#{XML.escape_attr(border[:diagonal][:color])}"/>)
            end
            @writer << '</diagonal>'
          end
          @writer << '</border>'
        end
      end
      @writer << XML.strip(<<-XML)
          </borders>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellStyleXfs>
          <cellXfs count="#{@styles.size}">
      XML
      @styles.each do |style|
        @writer << %(<xf numFmtId="#{style[:numFmtId] || 0}" fontId="#{style[:fontId] || 0}" fillId="#{style[:fillId] || 0}" borderId="#{style[:borderId] || 0}" xfId="0")
        unless XML.blank?(style[:h_align]) && XML.blank?(style[:v_align]) && XML.blank?(style[:word_wrap])
          @writer << %( applyAlignment="1")
        end
        unless XML.blank?(style[:numFmtId])
          @writer << %( applyNumberFormat="1")
        end
        unless XML.blank?(style[:h_align]) && XML.blank?(style[:v_align]) && XML.blank?(style[:word_wrap])
          @writer << %(><alignment#{XML.blank?(style[:h_align])? '' : %( horizontal="#{style[:h_align]}")}#{XML.blank?(style[:v_align])? '' : %( vertical="#{style[:v_align]}")}#{XML.blank?(style[:word_wrap])? '' : ' textWrap="1"'}/></xf>)
        else
          @writer << '/>'
        end
      end
      @writer << XML.strip(<<-XML)
          </cellXfs>
          <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
          </cellStyles>
          <dxfs count="0"/>
          <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
        </styleSheet>
      XML
    end

    def write_sst
      @writer.add_file "xl/sharedStrings.xml"
      @writer << XML.header
      @writer << %'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="#{@sst.references}" uniqueCount="#{@sst.size}">'
      @sst.each_key do |string|
        @writer << "<si><t>#{XML.escape_value string}</t></si>"
      end
      @writer << '</sst>'
    end

    def write_workbook_rels
      rid = String.new("rId0")
      @writer.add_file "xl/_rels/workbook.xml.rels"
      @writer << XML.header
      @writer << '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      @worksheets.each do |worksheet|
        @writer << %'<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet#{worksheet.id}.xml"/>'
      end
      @writer << %'<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      @writer << %'<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' unless @sst.empty?
      @writer << '</Relationships>'
    end

    def write_content_types
      @writer.add_file "[Content_Types].xml"
      @writer << XML.header
      @writer << XML.strip(<<-XML)
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      XML
      @writer << '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' unless @sst.empty?
      @worksheets.each do |worksheet|
        @writer << %'<Override PartName="/xl/worksheets/sheet#{worksheet.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      end
      @writer << '</Types>'
    end
  end
end
