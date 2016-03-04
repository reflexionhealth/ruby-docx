require_relative 'elements'
require_relative 'units'

module Docx
  module Editor
    include Units

    def self.new_document
      doc = Document.new
      doc.document.background.color = 'FFFFFF'
      doc.set_page_size(width: Inches * 8.5, height: Inches * 11)
      doc.set_page_margins(top: Inches * 1, bottom: Inches * 1, left: Inches * 1, right: Inches * 1)
      doc.set_page_numbering(start: 1)
      doc
    end

    class Document
      include Elements

      attr_accessor :filename
      attr_accessor :document
      attr_accessor :numbering
      attr_accessor :numbering_styles
      attr_accessor :styles
      attr_accessor :settings

      def initialize
        @document = W::Document.new
        @numbering = W::Numbering.new
        @numbering_styles = {}
      end

      def define_list_style(style_name, numbering)
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
        paragraph = Paragraph.new(document)
        document.content.push(paragraph.paragraph)
        paragraph
      end
    end

    class Paragraph
      attr_accessor :document
      attr_accessor :paragraph

      def initialize(document)
        @document = document
        # TODO: Should we always set underline for a new paragraph?
        @paragraph = W::Paragraph.new{properties: {run: {underline: 'none'}}}
      end

      def set_list_style(style_name, indent: 0)
        style = @document.numbering_styles[style_name]
        raise "unknown list style '#{style_name}'" if style.nil?

        props = @paragraph.properties
        props.numbering.indent.val = 0
        props.numbering.id = style[:definition].id
        # TODO: Is this correct source for this data? Check more example files.
        props.indent.left = style[:abstract].levels[indent].paragraph.indent.left
        props.indent.hanging = style[:abstract].levels[0].paragraph.indent.first_line
        # TODO: Should these be the default for paragraphs, or added as part of +set_list_style+?
        props.contextual_spacing.val = 1
        props
      end

      def set_font_size(size)
        run = @paragraph.properties.run
        run.font_size = size
        run.font_size_complex = size
        run
      end

      def add_text(text)
        props = @paragraph.properties.run
        run = W::TextRun.new({
          text: {space: 'preserve', text: text},
          properties: {
            font_size: props.font_size,
            font_size_complex: props.font_size_complex,
            right_to_left: 0
          }
        })
        @paragraph.content.push(run)
        run
      end
    end
  end
end

# include Docx::Units
# doc = Docx::Document.new
# doc.set_page_size
# doc.set_page_margins
