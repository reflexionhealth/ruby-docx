require_relative '../xmlish/tag'

module Docx
  module Elements
    module Typ
      Namespace = 'http://schemas.openxmlformats.org/package/2006/content-types'.freeze
      class Default < Xmlish::Tag
        type 'Default'
        namespace Namespace
        attribute :content, 'ContentType'
        attribute :ext, 'Extension'
      end
      class Override < Xmlish::Tag
        type 'Override'
        namespace Namespace
        attribute :content, 'ContentType'
        attribute :path, 'PartName'
      end
      class Types < Xmlish::Tag
        type 'Types'
        namespace Namespace
        tags :defaults, Default
        tags :overrides, Override
      end
    end

    module Rel
      Namespace = 'http://schemas.openxmlformats.org/package/2006/relationships'.freeze
      class Relationship < Xmlish::Tag
        type 'Relationship'
        namespace Namespace
        attribute :id, 'Id'
        attribute :type, 'Type'
        attribute :target, 'Target'
      end
      class Relationships < Xmlish::Tag
        type 'Relationships'
        namespace Namespace
        tags :rels, Relationship
      end
    end

    module W
      Namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'.freeze
      Tag = Xmlish::Tag

      # NOTE: When defining attributes, ensure they are defined in the same
      # order as they are serialized when downloading a .docx from Google Docs.
      def self.define(tag_type, tag_attributes)
        Class.new(Tag) do
          type tag_type
          namespace Namespace
          attributes *tag_attributes
        end
      end

      # Properties
      class RunFonts < Tag
        type 'rFonts'
        namespace Namespace
        attribute :ascii
        attribute :complex, 'cs'
        attribute :east_asia
        attribute :ansi, 'hAnsi'
      end
      class RunProperties < Tag
        type 'rPr'
        namespace Namespace
        tag :fonts, RunFonts
        tag :bold, W.define('b', [:val])
        tag :italic, W.define('i', [:val])
        tag :small_caps, W.define('smallCaps', [:val])
        tag :strike, W.define('strike', [:val])
        tag :color, W.define('color', [:val])
        tag :font_size, W.define('sz', [:val])
        tag :font_size_complex, W.define('szCs', [:val])
        tag :underline, W.define('u', [:val])
        tag :vertical_align, W.define('vertAlign', [:val])
        tag :right_to_left, W.define('rtl', [:val])
      end
      class NumberingProperties < Tag
        type 'numPr'
        namespace Namespace
        tag :indent, W.define('ilvl', [:val])
        tag :id, W.define('numId', [:val])
      end
      class ParagraphProperties < Tag
        type 'pPr'
        namespace Namespace
        tag :keep_next, W.define('keepNext', [:val])
        tag :keep_lines, W.define('keepLines', [:val])
        tag :widow_control, W.define('widowControl', [:val])
        tag :numbering, NumberingProperties
        tag :spacing, W.define('spacing', [:after, :before, :bottom, :line, :line_rule])
        tag :indent, W.define('ind', [:left, :right, :hanging, :first_line])
        tag :contextual_spacing, W.define('contextualSpacing', [:val])
        tag :justify, W.define('jc', [:val])
        tag :run, RunProperties
      end
      class PageSize < Tag
        type 'pgSz'
        namespace Namespace
        attribute :height, 'h'
        attribute :width, 'w'
      end
      class PageMargins < Tag
        type 'pgMar'
        namespace Namespace
        attributes :bottom, :top, :left, :right
        attributes :header, :footer, :gutter
      end
      class PageNumbering < Tag
        type 'pgNumType'
        namespace Namespace
        attribute :start
      end
      class SectionProperties < Tag
        type 'sectPr'
        namespace Namespace
        tag :page_size, PageSize
        tag :page_margins, PageMargins
        tag :page_numbering, PageNumbering
      end
      class Background < Tag
        namespace Namespace
        attribute :color
        content :tags, :tags
      end

      # Document
      class Text < Tag
        type 't'
        namespace Namespace
        attribute :space, prefix: 'xml'
        content :text, :text
      end
      class Run < Tag
        type 'r'
        namespace Namespace
        attribute :rev_id_deletion, 'rsidDel'
        attribute :rev_id_run, 'rsidR'
        attribute :rev_id_properties, 'rsidRPr'
        tag :properties, RunProperties
        tags :content, [Text]

        def initialize(*args)
          self.rev_id_deletion = '00000000'
          self.rev_id_run = '00000000'
          self.rev_id_properties = '00000000'
          super(*args)
        end
      end
      class Paragraph < Tag
        type 'p'
        namespace Namespace
        attribute :rev_id_paragraph, 'rsidR'
        attribute :rev_id_deletion, 'rsidDel'
        attribute :rev_id_properties, 'rsidP'
        attribute :rev_id_runs_default, 'rsidRDefault'
        attribute :rev_id_glyph_format, 'rsidRPr'
        tag :properties, ParagraphProperties
        tags :content, [Run]

        def initialize(*args)
          self.rev_id_paragraph = '00000000'
          self.rev_id_deletion = '00000000'
          self.rev_id_properties = '00000000'
          self.rev_id_runs_default = '00000000'
          self.rev_id_glyph_format = '00000000'
          super(*args)
        end
      end
      class Table < Tag
        type 'tbl'
        namespace Namespace
      end
      class Body < Tag
        namespace Namespace
        tags :content, [Paragraph, Table]
        tag :properties, SectionProperties
      end
      class Document < Tag
        namespace W::Namespace
        tag :background, W::Background
        tag :body, W::Body
      end

      # Fonts
      class Font < Tag
        namespace W::Namespace
        attribute :name
      end
      class FontTable < Tag
        type 'fonts'
        namespace W::Namespace
        tags :fonts, Font
      end

      # Numbering (Lists)
      class LevelDefinition < Tag
        type 'lvl'
        namespace W::Namespace
        attribute :level, 'ilvl'
        tag :start, W.define('start', [:val])
        tag :format, W.define('numFmt', [:val])
        tag :text, W.define('lvlText', [:val])
        tag :justify, W.define('lvlJc', [:val])
        tag :paragraph, ParagraphProperties
        tag :run, RunProperties
      end
      class AbstractNumberDefinition < Tag
        type 'abstractNum'
        namespace W::Namespace
        attribute :id, 'abstractNumId'
        tags :levels, LevelDefinition, max: 9
      end
      class NumberDefinition < Tag
        type 'num'
        namespace W::Namespace
        attribute :id, 'numId'
        tag :abstract_id, W.define('abstractNumId', [:val])
      end
      class Numbering < Tag
        namespace W::Namespace
        tags :abstract_definitions, AbstractNumberDefinition
        tags :definitions, NumberDefinition
      end

      # Settings
      class Compatibility < Tag
        type 'compat'
        namespace W::Namespace
        # NOTE: The primary schema reference I'm using doesn't list this element.
        # See http://www.datypic.com/sc/ooxml/s-wml.xsd.html.
        # According to the Microsoft Open XML SDK reference, this tag is always last.
        tag :setting, W.define('compatSetting', [:val, :name, :uri])
      end
      class Settings < Tag
        namespace W::Namespace
        tag :display_background_shapes, W.define('displayBackgroundShape', [:val])
        tag :default_tab_stop, W.define('defaultTabStop', [:val])
        tag :compatibility, Compatibility
      end

      # Styles
      class DefaultRun < Tag
        type 'rPrDefault'
        namespace W::Namespace
        tag :properties, RunProperties
      end
      class DefaultParagraph < Tag
        type 'pPrDefault'
        namespace W::Namespace
        tag :properties, ParagraphProperties
      end
      class DefaultProperties < Tag
        type 'docDefaults'
        namespace W::Namespace
        tag :run, DefaultRun
        tag :paragraph, DefaultParagraph
      end
      class Style < Tag
        namespace W::Namespace
        attribute :type
        attribute :id, 'styleId'
        attribute :default
        tag :name, W.define('name', [:val])
        tag :aliases, W.define('aliases', [:val])
        tag :based_on, W.define('basedOn', [:val])
        tag :next, W.define('next', [:val])
        tag :paragraph, ParagraphProperties
        tag :run, RunProperties
      end
      class Styles < Tag
        namespace W::Namespace
        tag :default, DefaultProperties
        tags :styles, Style
      end
    end
  end
end
