# frozen_string_literal: true
require 'test_helper'
require 'stringio'
require_relative '../../lib/rich_xlsx/worksheet.rb'

module RichXlsx
  class WorksheetTest < Minitest::Test
    def test_empty_worksheet
      io = StringIO.new
      ws = Worksheet.new(io)
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_add_row
      io = StringIO.new
      ws = Worksheet.new(io)
      ws << ['foo']
      ws.add_row ['bar']
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' \
          '<row r="1"><c r="A1" t="inlineStr"><is><t>foo</t></is></c></row>' \
          '<row r="2"><c r="A2" t="inlineStr"><is><t>bar</t></is></c></row>' \
        '</sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_add_row_with_sst_option
      io = StringIO.new
      mock_sst = { 'foo' => 0 }
      ws = Worksheet.new(io, :sst => mock_sst)
      ws << ['foo']
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' \
          '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' \
        '</sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_add_row_with_auto_format_option
      io = StringIO.new
      ws = Worksheet.new(io, :auto_format => true)
      ws << ['1.5']
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' \
          '<row r="1"><c r="A1" t="n"><v>1.5</v></c></row>' \
        '</sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_add_row_with_has_header_row_option
      io = StringIO.new
      ws = Worksheet.new(io, :has_header_row => true)
      ws << ['header']
      ws.add_row ['not header']
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' \
          '<row r="1"><c r="A1" s="3" t="inlineStr"><is><t>header</t></is></c></row>' \
          '<row r="2"><c r="A2" t="inlineStr"><is><t>not header</t></is></c></row>' \
        '</sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_add_columns_via_worksheet_options
      io = StringIO.new
      ws = Worksheet.new(io, { :columns => [ {}, {}, { :width_pixels => 42 } ] } )
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>' \
          '<col min="1" max="1"/>' \
          '<col min="2" max="2"/>' \
          '<col min="3" max="3" width="42" customWidth="1"/>' \
        '</cols>' \
        '<sheetData></sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_add_columns_via_worksheet_options_and_add_rows
      io = StringIO.new
      ws = Worksheet.new(io, { :columns => [ {}, {}, { :width_pixels => 42 } ] } )
      ws << ['foo']
      ws.add_row ['bar']
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>' \
          '<col min="1" max="1"/>' \
          '<col min="2" max="2"/>' \
          '<col min="3" max="3" width="42" customWidth="1"/>' \
        '</cols>' \
        '<sheetData>' \
          '<row r="1"><c r="A1" t="inlineStr"><is><t>foo</t></is></c></row>' \
          '<row r="2"><c r="A2" t="inlineStr"><is><t>bar</t></is></c></row>' \
        '</sheetData></worksheet>'
      assert_equal expected, io.string
    end

    def test_respond_to_id
      ws = Worksheet.new(StringIO.new, id: 1)
      assert_equal 1, ws.id
    end

    def test_respond_to_name
      ws = Worksheet.new(StringIO.new, name: 'test')
      assert_equal 'test', ws.name
    end
    
    def test_freeze_pane_and_xy_to_cell
      io = StringIO.new
      ws = Worksheet.new(io, { :columns => [ {}, {}, {}, { :width_pixels => 42 }, {} ] } )
      ws.freeze_pane(2, 3)
      5.times do |row_index|
        row = []
        5.times do |col_index|
          row << Worksheet.xy_to_cell(col_index + 1, row_index + 1)
        end
        ws << row
      end
      ws.close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" state="frozen"/></sheetView></sheetViews><cols>' \
          '<col min="1" max="1"/>' \
          '<col min="2" max="2"/>' \
          '<col min="3" max="3"/>' \
          '<col min="4" max="4" width="42" customWidth="1"/>' \
          '<col min="5" max="5"/>' \
        '</cols>' \
        '<sheetData>' \
          '<row r="1"><c r="A1" t="inlineStr"><is><t>A1</t></is></c><c r="B1" t="inlineStr"><is><t>B1</t></is></c><c r="C1" t="inlineStr"><is><t>C1</t></is></c><c r="D1" t="inlineStr"><is><t>D1</t></is></c><c r="E1" t="inlineStr"><is><t>E1</t></is></c></row>' \
          '<row r="2"><c r="A2" t="inlineStr"><is><t>A2</t></is></c><c r="B2" t="inlineStr"><is><t>B2</t></is></c><c r="C2" t="inlineStr"><is><t>C2</t></is></c><c r="D2" t="inlineStr"><is><t>D2</t></is></c><c r="E2" t="inlineStr"><is><t>E2</t></is></c></row>' \
          '<row r="3"><c r="A3" t="inlineStr"><is><t>A3</t></is></c><c r="B3" t="inlineStr"><is><t>B3</t></is></c><c r="C3" t="inlineStr"><is><t>C3</t></is></c><c r="D3" t="inlineStr"><is><t>D3</t></is></c><c r="E3" t="inlineStr"><is><t>E3</t></is></c></row>' \
          '<row r="4"><c r="A4" t="inlineStr"><is><t>A4</t></is></c><c r="B4" t="inlineStr"><is><t>B4</t></is></c><c r="C4" t="inlineStr"><is><t>C4</t></is></c><c r="D4" t="inlineStr"><is><t>D4</t></is></c><c r="E4" t="inlineStr"><is><t>E4</t></is></c></row>' \
          '<row r="5"><c r="A5" t="inlineStr"><is><t>A5</t></is></c><c r="B5" t="inlineStr"><is><t>B5</t></is></c><c r="C5" t="inlineStr"><is><t>C5</t></is></c><c r="D5" t="inlineStr"><is><t>D5</t></is></c><c r="E5" t="inlineStr"><is><t>E5</t></is></c></row>' \
        '</sheetData></worksheet>'
      assert_equal expected, io.string
    end
  end
end
