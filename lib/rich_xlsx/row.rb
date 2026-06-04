# frozen_string_literal: true
require "date"
require_relative "xml.rb"

module RichXlsx
  class Row

    ENCODING = Encoding.find('UTF-8')

    NUMBER_PATTERN = /\A-?[0-9]+(\.[0-9]+)?\z/.freeze
    # ISO 8601 yyyy-mm-dd
    DATE_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z/.freeze
    # ISO 8601 yyyy-mm-ddThh:mm:ss(.s)(Z|+hh:mm|-hh:mm)
    TIME_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,9})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})?\z/.freeze

    TRUE_STRING = 'true'.freeze
    FALSE_STRING = 'false'.freeze

    def initialize(row, rownum, options = {})
      @row = row
      @rownum = rownum
      @sst = options[:sst]
      @auto_format = options[:auto_format]
      @is_header = options[:is_header]
    end

    def to_xml
      column = String.new('A')
      xml = String.new(%Q{<row r="#{@rownum}">})

      if @is_header
        normal_style = ' s="3"'
        date_style   = ' s="4"'
        time_style   = ' s="5"'
      else
        normal_style = ''
        date_style   = ' s="1"'
        time_style   = ' s="2"'
      end

      @row.each do |value|
        cid = "#{column}#{@rownum}"
        column.next!
        style = nil
        if value.is_a?(Array)
          if value.size != 2
            raise "The [value, format] tuple should be an array with exactly 2 elements."
          end
          format_index = value[1]
          if (!format_index.is_a?(Numeric)) || (format_index < 0)
            raise "Invalid format index #{format_index}."
          end
          value = value[0]
          if format_index != 0
            style = " s=\"#{format_index.to_i}\""
          end
        end
        
        if @auto_format && value.is_a?(String)
          value = auto_format(value)
        end

        if style.nil?
          case value
          when Time, DateTime
            style = time_style
          when Date
            style = date_style
          else
            style = normal_style
          end
        end
        
        case value
        when Numeric
          xml << %Q{<c r="#{cid}"#{style} t="n"><v>#{value}</v></c>}
        when TrueClass, FalseClass
          xml << %Q{<c r="#{cid}"#{style} t="b"><v>#{value ? 1 : 0}</v></c>}
        when Time
          xml << %Q{<c r="#{cid}"#{style}><v>#{time_to_oa_date(value)}</v></c>}
        when DateTime
          xml << %Q{<c r="#{cid}"#{style}><v>#{datetime_to_oa_date(value)}</v></c>}
        when Date
          xml << %Q{<c r="#{cid}"#{style}><v>#{date_to_oa_date(value)}</v></c>}
        else
          value = value.to_s

          unless value.empty? # no xml output for for empty strings
            value = value.encode(ENCODING) if value.encoding != ENCODING

            if @sst
              xml << %Q{<c r="#{cid}"#{style} t="s"><v>#{@sst[value]}</v></c>}
            else
              xml << %Q{<c r="#{cid}"#{style} t="inlineStr"><is><t>#{XML.escape_value(value)}</t></is></c>}
            end
          end
        end
      end

      xml << '</row>'
    end

    private

    # Detects and casts numbers, date, time in text
    def auto_format(value)
      case value
      when TRUE_STRING
        true
      when FALSE_STRING
        false
      when NUMBER_PATTERN
        value.include?('.') ? value.to_f : value.to_i
      when DATE_PATTERN
        Date.parse(value) rescue value
      when TIME_PATTERN
        DateTime.parse(value) rescue value
      else
        value
      end
    end

    # Converts Time instance to OLE Automation Date
    def time_to_oa_date(time)
      # Local dates are stored as UTC by truncating the offset:
      # 1970-01-01 00:00:00 +0200 => 1970-01-01 00:00:00 UTC
      # This is done because SpreadsheetML is not timezone aware.
      (time.to_f + time.utc_offset) / 86400 + 25569
    end

    # Converts DateTime instance to OLE Automation Date
    if RUBY_ENGINE == 'ruby'
      def datetime_to_oa_date(date)
        _, jd, df, sf, of = date.marshal_dump
        jd - 2415019 + (df + of + sf / 1e9) / 86400
      end
    else
      def datetime_to_oa_date(date)
        date.jd - 2415019 + (date.hour * 3600 + date.sec + date.sec_fraction.to_f) / 86400
      end
    end

    # Converts Date instance to OLE Automation Date
    def date_to_oa_date(date)
      (date.jd - 2415019).to_f
    end
  end
end
