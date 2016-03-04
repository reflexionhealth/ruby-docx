require_relative 'tag'

module Docx
  module Elements
    module W
      Schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'.freeze

      def self.define(tag_name, tag_attributes)
        Class.new(Tag) do
          title(tag_name)
          schema(Schema)
          attributes(*tag_attributes)
        end
      end

      # Properties
      class RunProperties < Tag
        title 'rPr'
        schema Schema
        tag :bold, W.define('b', [:val])
        tag :italic, W.define('i', [:val])
        tag :font_size, W.define('sz', [:val])
        tag :font_size_complex, W.define('szCs', [:val])
        tag :underline, W.define('u', [:val, :color])
        tag :right_to_left, W.define('rtl', [:val])
      end
      class NumberingProperties < Tag
        title 'numPr'
        schema Schema
        tag :indent, W.define('ilvl', [:val])
        tag :id, W.define('numId', [:val])
      class ParagraphProperties < Tag
        title 'pPr'
        schema Schema
        tag :numbering, NumberingProperties
        tag :indent, W.define('ind', [:left, :right, :hanging, :first_line])
        tag :contextual_spacing, W.define('contextual_spacing', [:val])
        tag :run, RunProperties
      end
      class PageSize < Tag
        title 'pgSz'
        schema Schema
        attribute :width, 'w'
        attribute :height, 'h'
      end
      class PageMargins < Tag
        title 'pgMar'
        schema Schema
        attributes :top, :right, :bottom, :left
        attributes :header, :footer, :gutter
      end
      class PageNumbering < Tag
        title 'pgNumType'
        schema Schema
        attribute :start
      end
      class SectionProperties < Tag
        title 'sectPr'
        schema Schema
        tag :page_size, PageSize
        tag :page_margins, PageMargins
        tag :page_numbering, PageNumbering
      end
      class Background < Tag
        schema Schema
        attribute :color
        content :tags, :tags
      end

      # Content
      class Text < Tag
        title 'text'
        schema Schema
        attribute :space
        content :text, :text
      end

      class TextRun < Tag
        title 'r'
        schema Schema
        tag :properties, RunProperties
        tags :content, [Text]
      end
      class Paragraph < Tag
        title 'p'
        schema Schema
        tag :properties, ParagraphProperties
        tags :content, [TextRun]
      end
      class Table < Tag; title 'tbl'; schema Schema; end
      class Body < Tag
        schema Schema
        tags :content, [Paragraph, Table]
        tag :properties, SectionProperties
      end

      # Document
      class Document < Tag
        schema W::Schema
        tag :background, W::Background
        tag :body, W::Body
      end

      # (List) Numbering
      class LevelDefinition < Tag
        title 'lvl'
        schema W::Schema
        attribute :level, 'ilvl'
        tag :start, W.define('start', [:val])
        tag :format, W.define('numFmt', [:val])
        tag :text, W.define('lvlText', [:val])
        tag :justify, W.define('lvlJc', [:val])
        tag :paragraph, ParagraphProperties
        tag :run, RunProperties
      end
      class AbstractNumberDefinition < Tag
        title 'abstractNum'
        schema W::Schema
        attribute :id, 'abstractNumId'
        tags :levels, LevelDefinition, max: 9
      end
      class NumberDefinition < Tag
        title 'num'
        schema W::Schema
        attribute :id, 'numId'
        tag :abstract_id, W.define('abstractNumId', [:val])
      end
      class Numbering < Tag
        schema W::Schema
        tags :abstract_definitions, AbstractNumberDefinition
        tags :definitions, NumberDefinition
      end
    end
  end
end
