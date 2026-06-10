# frozen_string_literal: true
require_relative "rich_xlsx/version.rb"

module RichXlsx
  EXCEL_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
end

require_relative "rich_xlsx/workbook.rb"
require_relative "rich_xlsx/worksheet.rb"
require_relative "rich_xlsx/columns.rb"
require_relative "rich_xlsx/row.rb"
