require_relative '../xmlish/encoding'
require_relative 'elements'
require_relative 'styles'
require_relative 'units'

module Docx
  module Editor
    include Elements
    include Units

    Namespaces = Hash[
      'mc' => 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'o' => 'urn:schemas-microsoft-com:office:office',
      'r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'm' => 'http://schemas.openxmlformats.org/officeDocument/2006/math',
      'v' => 'urn:schemas-microsoft-com:vml',
      'wp' => 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
      'w10' => 'urn:schemas-microsoft-com:office:word',
      'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'wne' => 'http://schemas.microsoft.com/office/word/2006/wordml',
      'sl' => 'http://schemas.openxmlformats.org/schemaLibrary/2006/main',
      'a' => 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'pic' => 'http://schemas.openxmlformats.org/drawingml/2006/picture',
      'c' => 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'lc' => 'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas',
      'dgm' => 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
      'wps' => 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
      'wpg' => 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    ].freeze

    def self.new_document
      doc = Document.new
      doc.document.background.color = 'FFFFFF'
      doc.set_page_size(width: Inches * 8.5, height: Inches * 11)
      doc.set_page_margins(top: Inches * 1, bottom: Inches * 1, left: Inches * 1, right: Inches * 1)
      doc.set_page_numbering(start: 1)
      doc.font_table.fonts.push(Elements::W::Font.new(name: 'Arial'))
      doc.settings.display_background_shapes.val = 1
      doc.settings.default_tab_stop.val = Units::Inches * 0.5
      doc.styles.default = Styles.default
      doc.styles.styles = [
        Styles.paragraph, Styles.table,
        Styles.h1, Styles.h2, Styles.h3,
        Styles.h4, Styles.h5, Styles.h6,
        Styles.type, Styles.subtitle
      ]
      doc
    end

    class Document
      include Elements

      attr_accessor :filename
      attr_accessor :document
      attr_accessor :font_table
      attr_accessor :numbering
      attr_accessor :numbering_styles
      attr_accessor :settings
      attr_accessor :styles

      def initialize
        @document = W::Document.new
        @font_table = W::FontTable.new
        @numbering = W::Numbering.new
        @numbering_styles = {}

        compat = {val: 14, name: 'compatibilityMode', uri: 'http://schemas.microsoft.com/office/word'}
        @settings = W::Settings.new({compatibility: {setting: compat}})
        @styles = W::Styles.new
      end

      def define_list_style(style_name, numbering_style)
        style = numbering_style
        style.id = @numbering.abstract_definitions.length + 1
        @numbering.abstract_definitions.push(style)

        defn = W::NumberDefinition.new({abstract_id: {val: style.id}})
        defn.id = @numbering.definitions.length + 1
        @numbering.definitions.push(defn)

        @numbering_styles[style_name] = {abstract: style, definition: defn}
      end

      def set_page_size(width: nil, height: nil)
        page_size = @document.body.properties.page_size
        page_size.width = width if width
        page_size.height = height if height
        page_size
      end

      def set_page_margins(top: nil, right: nil, bottom: nil, left: nil)
        page_margins = @document.body.properties.page_margins
        page_margins.top = top if top
        page_margins.right = right if right
        page_margins.bottom = bottom if bottom
        page_margins.left = left if left
        page_margins
      end

      def set_page_numbering(start: nil)
        page_numbering = @document.body.properties.page_numbering
        page_numbering.start = start if start
        page_numbering
      end

      def add_paragraph
        paragraph = Paragraph.new(self)
        @document.body.content.push(paragraph.paragraph)
        paragraph
      end
    end

    class Paragraph
      include Elements

      attr_accessor :document
      attr_accessor :paragraph

      def initialize(editor_doc)
        @document = editor_doc
        @paragraph = W::Paragraph.new
        # TODO: Should we always set underline for a new paragraph?
        @paragraph.properties.run.underline.val = 'none'
        self.set_font_size(Units::Points * 8)
      end

      def set_list_style(style_name, indent: 0)
        style = @document.numbering_styles[style_name]
        raise KeyError, "unknown list style '#{style_name}'" if style.nil?
        max_indent = style[:abstract].levels.count - 1
        if indent > max_indent
          raise IndexError, "can't use indent #{indent} with list style '#{style_name}' (max #{max_indent})"
        end

        props = @paragraph.properties
        props.numbering.indent.val = 0
        props.numbering.id.val = style[:definition].id
        # TODO: Is this correct source for this data? Check more example files.
        props.indent.left = style[:abstract].levels[indent].paragraph.indent.left
        props.indent.hanging = Units::Inches * 0.5
        # TODO: Should these be the default for paragraphs, or added as part of +set_list_style+?
        props.contextual_spacing.val = 1
        props
      end

      def set_font_size(size)
        run = @paragraph.properties.run
        run.font_size.val = size
        run.font_size_complex.val = size
        run
      end

      def add_text(text)
        props = @paragraph.properties.run
        run = W::Run.new({
          content: [W::Text.new({space: 'preserve', text: text})],
          properties: {
            font_size: props.font_size,
            font_size_complex: props.font_size_complex,
            right_to_left: {val: 0}
          }
        })
        @paragraph.content.push(run)
        run
      end
    end
  end
end
