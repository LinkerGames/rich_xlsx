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
      @number_formats = [
        "yyyy\\-mm\\-dd",
        "yyyy\\-mm\\-dd hh:mm:ss"
      ]
      @fonts = [
        {
          size: @options.fetch(:font, {}).fetch(:size, 12),
          name: @options.fetch(:font, {}).fetch(:name, 'Calibri'),
        },
        {
          size: @options.fetch(:font, {}).fetch(:size, 12),
          name: @options.fetch(:font, {}).fetch(:name, 'Calibri'),
          bold: true
        }
      ]
      @fills = [
        {
          patternType: 'none'
        },
        {
          patternType: 'gray125'
        }
      ]
      @borders = [{}]
      
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
          <numFmts count="#{@number_formats.size}">#{@number_formats.map.with_index{|fmt, index| %(<numFmt numFmtId="#{164 + index}" formatCode="#{XML.escape_attr(fmt)}"/>)}.join("\n")}</numFmts>
          <fonts count="#{@fonts.size}">#{@fonts.map{|font| %(<font>#{font[:bold]? '<b/>' : ''}#{font[:italic]? '<i/>' : ''}#{font[:underline]? '<u/>' : ''}#{font[:strike]? '<strike/>' : ''}#{XML.blank?(font[:color])? '' : %(<color rgb="#{XML.escape_attr(font[:color])}"/>)}#{XML.blank?(font[:size])? '' : %(<sz val="#{XML.escape_attr(font[:size].to_s)}"/>)}#{XML.blank?(font[:name])? '' : %(<name val="#{XML.escape_attr(font[:name])}"/>)}</font>)}}</fonts>
          <fills count="#{@fills.size}">#{@fills.map{|fill| %(<fill><patternFill patternType="#{fill[:patternType]}"#{(fill[:patternType]!='solid')? "/>" : %(>#{XML.blank?(fill[:fgColor])?'':%(<fgColor rgb="#{XML.escape_attr(fill[:fgColor])}"/>)}#{XML.blank?(fill[:bgColor])? '' : %(<fgColor rgb="#{XML.escape_attr(fill[:bgColor])}"/>)}</patternFill>)}</fill>)}.join("\n")}</fills>
          <borders count="#{@borders.size}">#{@borders.map{|border| %(<border#{XML.blank?(border)? '/>' : %(>#{VALID_OUTLINE_BORDER_KEYS.map{|key| %(#{XML.blank?(border[key])? '' : %(<#{key}#{XML.blank?(border[key][:style])? '' : %( style="#{border[key][:style]}")}>#{XML.blank?(border[key][:color])? '' : %(<color rgb="#{XML.escape_attr(border[key][:color])}"/>)}</#{key}>)})}}#{XML.blank?(border[:diagonal])? '' : %(<diagonal#{XML.blank?(border[:diagonal][:diagonalUp])? '' : %( diagonalUp="#{(border[:diagonal][:diagonalUp])? '1' : '0'}")}#{XML.blank?(border[:diagonal][:diagonalDown])? '' : %( diagonalDown="#{(border[:diagonal][:diagonalDown])? '1' : '0'}")}>#{XML.blank?(border[:diagonal][:color])? '' : %(<color rgb="#{XML.escape_attr(border[:diagonal][:color])}"/>)}</diagonal>)}</border>)})}}</borders>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellStyleXfs>
          <cellXfs count="6">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
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
