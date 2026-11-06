# frozen_string_literal: true
require_relative "xml.rb"

module RichXlsx
  # Writes the parts that make up the cell comments (notes) of one worksheet:
  # - xl/commentsN.xml                       - the comment texts and authors
  # - xl/drawings/vmlDrawingN.vml            - the legacy VML shapes Excel needs to display the notes
  # - xl/worksheets/_rels/sheetN.xml.rels    - links the worksheet to both parts
  # The worksheet itself references the VML drawing with <legacyDrawing r:id="rId1"/>.
  class Comments
    COMMENTS_RELATIONSHIP_ID = 'rId1'
    VML_DRAWING_RELATIONSHIP_ID = 'rId2'

    # Excel numbers VML shapes in blocks of 1024 ids per <o:idmap> entry.
    SHAPES_PER_ID_BLOCK = 1023

    # Approximate size of a comment box: it spans this many columns to the right of the cell
    # and grows one line per this many characters.
    BOX_COLUMNS = 3
    CHARS_PER_LINE = 50
    LINE_HEIGHT_PT = 12
    BOX_WIDTH_PT = 200

    attr_reader :worksheet_id

    # comments:         Array of {ref:, x:, y:, text:, author:} as collected by Worksheet#add_comment
    #                   (x and y are 1-based).
    # first_id_block:   first VML id block for this worksheet; blocks must be unique within the workbook.
    # default_author:   used for comments added without an author.
    def initialize(worksheet_id, comments, first_id_block, default_author)
      @worksheet_id = worksheet_id
      @comments = comments
      @first_id_block = first_id_block
      @default_author = default_author.to_s
    end

    def self.id_blocks_needed(comment_count)
      (comment_count.to_f / SHAPES_PER_ID_BLOCK).ceil
    end

    def id_blocks
      self.class.id_blocks_needed(@comments.size)
    end

    def comments_path
      "xl/comments#{@worksheet_id}.xml"
    end

    def vml_drawing_path
      "xl/drawings/vmlDrawing#{@worksheet_id}.vml"
    end

    def worksheet_rels_path
      "xl/worksheets/_rels/sheet#{@worksheet_id}.xml.rels"
    end

    def write(writer)
      write_comments(writer)
      write_vml_drawing(writer)
      write_worksheet_rels(writer)
    end

    private

    def authors
      @authors ||= @comments.map { |comment| author_of(comment) }.uniq
    end

    def author_of(comment)
      comment[:author].nil? ? @default_author : comment[:author].to_s
    end

    def write_comments(writer)
      writer.add_file comments_path
      writer << XML.header
      writer << '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>'
      authors.each do |author|
        writer << "<author>#{XML.escape_value(author)}</author>"
      end
      writer << '</authors><commentList>'
      @comments.each do |comment|
        writer << %(<comment ref="#{comment[:ref]}" authorId="#{authors.index(author_of(comment))}"><text><r>)
        writer << '<rPr><sz val="9"/><rFont val="Tahoma"/></rPr>'
        writer << %(<t xml:space="preserve">#{XML.escape_value(comment[:text])}</t>)
        writer << '</r></text></comment>'
      end
      writer << '</commentList></comments>'
    end

    def write_vml_drawing(writer)
      writer.add_file vml_drawing_path
      blocks = (@first_id_block...(@first_id_block + id_blocks)).to_a
      writer << '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
      writer << %(<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="#{blocks.join(',')}"/></o:shapelayout>)
      writer << '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">'
      writer << '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
      @comments.each_with_index do |comment, index|
        shape_id = (blocks[index / SHAPES_PER_ID_BLOCK] * 1024) + (index % SHAPES_PER_ID_BLOCK) + 1
        # VML rows and columns are 0-based.
        column = comment[:x] - 1
        row = comment[:y] - 1
        lines = text_lines(comment[:text])
        anchor = [column + 1, 15, row, 10, column + 1 + BOX_COLUMNS, 15, row + lines + 1, 4].join(', ')
        writer << %(<v:shape id="_x0000_s#{shape_id}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:2pt;width:#{BOX_WIDTH_PT}pt;height:#{(lines + 1) * LINE_HEIGHT_PT}pt;z-index:#{index + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">)
        writer << '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>'
        writer << '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>'
        writer << '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
        writer << %(<x:Anchor>#{anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>#{row}</x:Row><x:Column>#{column}</x:Column>)
        writer << '</x:ClientData></v:shape>'
      end
      writer << '</xml>'
    end

    def write_worksheet_rels(writer)
      writer.add_file worksheet_rels_path
      writer << XML.header
      writer << '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      writer << %(<Relationship Id="#{COMMENTS_RELATIONSHIP_ID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments#{@worksheet_id}.xml"/>)
      writer << %(<Relationship Id="#{VML_DRAWING_RELATIONSHIP_ID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing#{@worksheet_id}.vml"/>)
      writer << '</Relationships>'
    end

    def text_lines(text)
      text.split("\n", -1).sum { |line| [(line.length.to_f / CHARS_PER_LINE).ceil, 1].max }
    end
  end
end
