# frozen_string_literal: true
require 'test_helper'
require 'stringio'
require 'tempfile'
require_relative '../../lib/rich_xlsx/workbook.rb'

module RichXlsx
  class WorksheetTest < Minitest::Test

    class SpyWriter < RichXlsx::ZipKitWriter
      def initialize
        @paths_to_file_contents = {}
        @current = nil
      end

      def <<(data)
        @paths_to_file_contents[@current] << data
      end

      def add_file(path)
        @current = path
        @paths_to_file_contents[@current] = String.new
      end

      def [](key)
        @paths_to_file_contents.fetch(key)
      end

      def close
      end
    end

    def test_workbook_from_path
      tempfile = Tempfile.new('rich_xlsx')
      Workbook.open(tempfile.path) {}
      refute_equal 0, tempfile.size
    ensure
      tempfile.close! if tempfile
    end

    def test_workbook_from_io
      tempfile = Tempfile.new('rich_xlsx')
      Workbook.open(tempfile) {}
      refute_equal 0, tempfile.size
    ensure
      tempfile.close! if tempfile
    end

    def test_empty_workbook
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) {}
      expected = {
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets></sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_workbook_with_sheet
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet
      end
      expected = {
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
            '<sheetData></sheetData>' \
          '</worksheet>',
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_deprecated_add_workbook_with_block
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        silence_warnings do
          wb.add_worksheet {}
        end
      end
      expected = {
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
            '<sheetData></sheetData>' \
          '</worksheet>',
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_workbook_with_sheet_without_block
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        ws = wb.add_worksheet
        ws << ['foo']
        ws.close
      end
      expected = {
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
            '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>foo</t></is></c></row></sheetData>' \
          '</worksheet>',
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_workbook_with_sst
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet(nil, use_shared_strings: true) do |ws|
          ws << ['foo']
        end
      end
      expected = {
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
            '<sheetData>' \
              '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' \
            '</sheetData>' \
          '</worksheet>',
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/sharedStrings.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">' \
            '<si><t>foo</t></si>' \
          '</sst>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_root_relations
      iow_spy = io_wrapper_spy
      Workbook.new(iow_spy).close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' \
        '</Relationships>'
      actual = iow_spy['_rels/.rels']
      assert_equal expected, actual
    end

    def test_content_types
      iow_spy = io_wrapper_spy
      Workbook.new(iow_spy).close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' \
          '<Default Extension="xml" ContentType="application/xml"/>' \
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' \
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' \
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' \
        '</Types>'
      actual = iow_spy['[Content_Types].xml']
      assert_equal expected, actual
    end

    def test_write_multiple_worksheets
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet
        wb.write_worksheet
      end

      expected = {
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
              '<sheet name="Sheet2" sheetId="2" r:id="rId2"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' \
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>',
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>',
        'xl/worksheets/sheet2.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_must_write_sequentially
      iow_spy1 = io_wrapper_spy

      Workbook.open(iow_spy1) do |wb|
        wb.add_worksheet.tap { |ws| ws.close }
        wb.add_worksheet.tap { |ws| ws.close }
      end

      iow_spy2 = io_wrapper_spy
      assert_raises(RichXlsx::Error) do
        Workbook.open(iow_spy2) do |wb|
          wb.add_worksheet
          wb.add_worksheet # adding a second worksheet without closing
        end
      end
    end

    def test_write_named_worksheet
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet('foo')
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
          '<workbookPr date1904="false"/>' \
          '<sheets>' \
            '<sheet name="foo" sheetId="1" r:id="rId1"/>' \
          '</sheets>' \
        '</workbook>'
      actual = iow_spy['xl/workbook.xml']
      assert_equal expected, actual
    end

    def test_write_unnamed_worksheet_with_options
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet(:use_shared_strings => true)
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
          '<workbookPr date1904="false"/>' \
          '<sheets>' \
            '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
          '</sheets>' \
        '</workbook>'
      actual = iow_spy['xl/workbook.xml']
      assert_equal expected, actual
    end

    def test_worksheet_name_as_option
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |workbook|
        workbook.write_worksheet(name: "foo")
      end
      expected = '<sheet name="foo" sheetId="1" r:id="rId1"/>'
      actual = iow_spy['xl/workbook.xml'][/<sheet [^>]+>/]
      assert_equal expected, actual
    end

    def test_add_columns_via_workbook_options
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy, { :columns => [ {}, {}, { :width_pixels => 42 } ] } ) do |wb|
        wb.write_worksheet {}
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>' \
          '<col min="1" max="1"/>' \
          '<col min="2" max="2"/>' \
          '<col min="3" max="3" width="42" customWidth="1"/>' \
        '</cols>' \
        '<sheetData></sheetData></worksheet>'

      actual = iow_spy['xl/worksheets/sheet1.xml']
      assert_equal expected, actual
    end

    def test_add_columns_via_workbook_options_and_add_rows
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy, { :columns => [ {}, {}, { :width_pixels => 42 } ] } ) do |wb|
        wb.write_worksheet do |ws|
          ws << ['foo']
          ws.add_row ['bar']
        end
      end

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

      actual = iow_spy['xl/worksheets/sheet1.xml']
      assert_equal expected, actual
    end

    def test_styles_content
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) {}
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
          '<numFmts count="2">' \
            '<numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/>' \
            '<numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd hh:mm:ss"/>' \
          '</numFmts>' \
          '<fonts count="2">' \
            '<font>' \
              '<sz val="11"/>' \
              '<name val="Calibri"/>' \
            '</font>' \
            '<font>' \
              '<b/>' \
              '<sz val="11"/>' \
              '<name val="Calibri"/>' \
            '</font>' \
          '</fonts>' \
          '<fills count="2">' \
            '<fill>' \
              '<patternFill patternType="none"/>' \
            '</fill>' \
            '<fill>' \
              '<patternFill patternType="gray125"/>' \
            '</fill>' \
          '</fills>' \
          '<borders count="1">' \
            '<border/>' \
          '</borders>' \
          '<cellStyleXfs count="1">' \
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' \
          '</cellStyleXfs>' \
          '<cellXfs count="6">' \
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' \
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' \
            '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' \
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1">' \
              '<alignment horizontal="center" vertical="center"/>' \
            '</xf>' \
            '<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">' \
              '<alignment horizontal="center" vertical="center"/>' \
            '</xf>' \
            '<xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">' \
              '<alignment horizontal="center" vertical="center"/>' \
            '</xf>' \
          '</cellXfs>' \
          '<cellStyles count="1">' \
            '<cellStyle name="Normal" xfId="0" builtinId="0"/>' \
          '</cellStyles>' \
          '<dxfs count="0"/>' \
          '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>' \
        '</styleSheet>'
      actual = iow_spy['xl/styles.xml']
      assert_equal expected, actual
    end

    def test_custom_font_size
      iow_spy = io_wrapper_spy
      font_options = { :size => 23 }
      Workbook.open(iow_spy, :font => font_options) {}
      expected = '<sz val="23"/>'
      actual = iow_spy['xl/styles.xml'][/<sz [^>]+>/]
      assert_equal expected, actual
    end

    def test_custom_font_name
      iow_spy = io_wrapper_spy
      font_options = { :name => 'Comic Sans MS' }
      Workbook.open(iow_spy, :font => font_options) {}
      expected = '<name val="Comic Sans MS"/>'
      actual = iow_spy['xl/styles.xml'][/<name [^>]+>/]
      assert_equal expected, actual
    end

    def test_many_things
      iow_spy = io_wrapper_spy
      default_font_options = {
        name: 'Comic Sans MS', size: 16
      }
      Workbook.open(iow_spy, font: default_font_options) do |wb|
        previous_time_style = wb.add_style({
          format: "hh:mm:ss.000",
          font: {
            name: 'Times New Roman',
            size: 14,
            strike: true,
            italic: true,
            color: 'FFAAAAAA'
          },
          fill: {
            patternType: 'darkGrid',
            fgColor: 'FFFFEEEE',
            bgColor: 'FFEEEEFF'
          },
          border: {
            left: {
              style: 'dashed',
              color: 'FFFFAAA1'
            },
            right: {
              style: 'dotted',
              color: 'FFFFAAA2'
            },
            top: {
              style: 'double',
              color: 'FFFFAAA3'
            },
            bottom: {
              style: 'hair',
              color: 'FFFFAAA4'
            },
            diagonal: {
              diagonalUp: true,
              diagonalDown: true,
              style: 'mediumDashed',
              color: 'FFFFAAA5'
            },
            h_align: 'right',
            v_align: 'center'
          }
        })
        current_time_style = wb.add_style({
          format: "hh:mm:ss.000",
          font: {
            name: 'Comic Sans MS',
            size: 14,
            bold: true,
            color: 'FF11FF11'
          },
          fill: {
            patternType: 'solid',
            fgColor: 'FF118811'
          },
          border: {
            left: {
              style: 'dashDot',
              color: 'FFAAFFA1'
            },
            right: {
              style: 'mediumDashDot',
              color: 'FFAAFFA2'
            },
            top: {
              style: 'dashDotDot',
              color: 'FFAAFFA3'
            },
            bottom: {
              style: 'mediumDashDotDot',
              color: 'FFAAFFA4'
            }
          },
          h_align: 'center',
          v_align: 'center'
        })
        position_style = wb.add_style({
          font: {
            name: 'Times New Roman',
            size: 18,
            bold: true,
            color: 'FFFFFFFF'
          },
          fill: {
            patternType: 'solid',
            fgColor: 'FF000000'
          },
          h_align: 'center',
          v_align: 'center',
          word_wrap: true
        })
        currency_style = wb.add_style({
          format: "€ 0.00",
          font: {
            bold: true
          },
          border: {
            left: {
              style: 'thick'
            },
            right: {
              style: 'thick'
            },
            top: {
              style: 'thick'
            },
            bottom: {
              style: 'thick'
            }
          },
          h_align: 'right',
          v_align: 'top'
        })
        wb.write_worksheet(name: 'Lap times', auto_format: true, has_header_row: true, columns: [ {width_pixels: 12}, {width_pixels: 20}, {width_pixels: 20}, {width_pixels: 15}, {width_pixels: 20} ]) do |ws|
          ws << ['Driver name', 'Previous record', 'Current record', 'Position', 'Award']
          one_minute = Time.at(60)
          one_minute = one_minute - one_minute.utc_offset
          ws << ['Maximus', [one_minute, previous_time_style], [one_minute - 3.056, current_time_style], [1, position_style], [12345.67, currency_style]]
          ws.merge('A3','E5')
          ws << [['FINISH', position_style]]
        end
      end
      style_xml = XML.strip(iow_spy['xl/styles.xml'])
      expected_style_xml = XML.strip(<<-XML)
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <numFmts count="4">
            <numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/>
            <numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd hh:mm:ss"/>
            <numFmt numFmtId="166" formatCode="hh:mm:ss.000"/>
            <numFmt numFmtId="167" formatCode="€ 0.00"/>
          </numFmts>
          <fonts count="6">
            <font>
              <sz val="16"/>
              <name val="Comic Sans MS"/>
            </font>
            <font>
              <b/>
              <sz val="16"/>
              <name val="Comic Sans MS"/>
            </font>
            <font>
              <i/>
              <strike/>
              <color rgb="FFAAAAAA"/>
              <sz val="14"/>
              <name val="Times New Roman"/>
            </font>
            <font>
              <b/>
              <color rgb="FF11FF11"/>
              <sz val="14"/>
              <name val="Comic Sans MS"/>
            </font>
            <font>
              <b/>
              <color rgb="FFFFFFFF"/>
              <sz val="18"/>
              <name val="Times New Roman"/>
            </font>
            <font>
              <b/>
            </font>
          </fonts>
          <fills count="5">
            <fill>
              <patternFill patternType="none"/>
            </fill>
            <fill>
              <patternFill patternType="gray125"/>
            </fill>
            <fill>
              <patternFill patternType="darkGrid">
                <fgColor rgb="FFFFEEEE"/>
                <bgColor rgb="FFEEEEFF"/>
              </patternFill>
            </fill>
            <fill>
              <patternFill patternType="solid">
                <fgColor rgb="FF118811"/>
              </patternFill>
            </fill>
            <fill>
              <patternFill patternType="solid">
                <fgColor rgb="FF000000"/>
              </patternFill>
            </fill>
          </fills>
          <borders count="4">
            <border/>
            <border diagonalUp="1" diagonalDown="1">
              <left style="dashed">
                <color rgb="FFFFAAA1"/>
              </left>
              <right style="dotted">
                <color rgb="FFFFAAA2"/>
              </right>
              <top style="double">
                <color rgb="FFFFAAA3"/>
              </top>
              <bottom style="hair">
                <color rgb="FFFFAAA4"/>
              </bottom>
              <diagonal style="mediumDashed">
                <color rgb="FFFFAAA5"/>
              </diagonal>
            </border>
            <border>
              <left style="dashDot">
                <color rgb="FFAAFFA1"/>
              </left>
              <right style="mediumDashDot">
                <color rgb="FFAAFFA2"/>
              </right>
              <top style="dashDotDot">
                <color rgb="FFAAFFA3"/>
              </top>
              <bottom style="mediumDashDotDot">
                <color rgb="FFAAFFA4"/>
              </bottom>
            </border>
            <border>
              <left style="thick">
              </left>
              <right style="thick">
              </right>
              <top style="thick">
              </top>
              <bottom style="thick">
              </bottom>
            </border>
          </borders>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellStyleXfs>
          <cellXfs count="10">
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
            <xf numFmtId="166" fontId="2" fillId="2" borderId="1" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="166" fontId="3" fillId="3" borderId="2" xfId="0" applyAlignment="1" applyNumberFormat="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="4" fillId="4" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="center" vertical="center" textWrap="1"/>
            </xf>
            <xf numFmtId="167" fontId="5" fillId="0" borderId="3" xfId="0" applyAlignment="1" applyNumberFormat="1">
              <alignment horizontal="right" vertical="top"/>
            </xf>
          </cellXfs>
          <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
          </cellStyles>
          <dxfs count="0"/>
          <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
        </styleSheet>
      XML
      assert_equal expected_style_xml, style_xml
      worksheet_xml = XML.strip(iow_spy['xl/worksheets/sheet1.xml'])
      expected_worksheet_xml = XML.strip(<<-XML)
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cols>
            <col min="1" max="1" width="12" customWidth="1"/>
            <col min="2" max="2" width="20" customWidth="1"/>
            <col min="3" max="3" width="20" customWidth="1"/>
            <col min="4" max="4" width="15" customWidth="1"/>
            <col min="5" max="5" width="20" customWidth="1"/>
          </cols>
          <sheetData>
            <row r="1">
              <c r="A1" s="3" t="inlineStr">
                <is>
                  <t>Driver name</t>
                </is>
              </c>
              <c r="B1" s="3" t="inlineStr">
                <is>
                  <t>Previous record</t>
                </is>
              </c>
              <c r="C1" s="3" t="inlineStr">
                <is>
                  <t>Current record</t>
                </is>
              </c>
              <c r="D1" s="3" t="inlineStr">
                <is>
                  <t>Position</t>
                </is>
              </c>
              <c r="E1" s="3" t="inlineStr">
                <is>
                  <t>Award</t>
                </is>
              </c>
            </row>
            <row r="2">
              <c r="A2" t="inlineStr">
                <is>
                  <t>Maximus</t>
                </is>
              </c>
              <c r="B2" s="6">
                <v>25569.000694444443</v>
              </c>
              <c r="C2" s="7">
                <v>25569.000659074074</v>
              </c>
              <c r="D2" s="8" t="n">
                <v>1</v>
              </c>
              <c r="E2" s="9" t="n">
                <v>12345.67</v>
              </c>
            </row>
            <row r="3">
              <c r="A3" s="8" t="inlineStr">
                <is>
                  <t>FINISH</t>
                </is>
              </c>
            </row>
          </sheetData>
          <mergeCells>
            <mergeCell ref="A3:E5"/>
          </mergeCells>
        </worksheet>
      XML
      assert_equal expected_worksheet_xml, worksheet_xml
    end
    
    def test_tempfile_is_not_closed
      tempfile = Tempfile.new('workbook')
      Workbook.open(tempfile) {}
      assert_equal false, tempfile.closed?
    ensure
      tempfile && tempfile.close!
    end

    private

    def io_wrapper_spy
      SpyWriter.new
    end

    def silence_warnings
      old_verbose, $VERBOSE = $VERBOSE, nil
      yield
    ensure
      $VERBOSE = old_verbose
    end

  end
end
