require_relative 'xml/encoding'
require_relative 'elements'
require_relative 'styles'
require_relative 'units'
require 'zipruby'
require 'tmpdir'

module Docx
  module Editor
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
      doc.settings.default_tab_stop.val = Inches * 0.5
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

      def save_as(filepath)
        @filename = filepath
        self.save
      end

      def save
        raise "cannot save a document without a filename; use 'save_as' instead" if @filename.nil?

        # build content types
        wpmlns = 'application/vnd.openxmlformats-officedocument.wordprocessingml'
        types = Typ::Types.new({
          defaults: [
            Typ::Default.new({content: 'application/xml', ext: 'xml'}),
            Typ::Default.new({content: 'application/vnd.openxmlformats-package.relationships+xml', ext: 'rels'})
          ],
          overrides: [
            Typ::Override.new({content: "#{wpmlns}.settings+xml", path: '/word/settings.xml'}),
            Typ::Override.new({content: "#{wpmlns}.styles+xml", path: '/word/styles.xml'}),
            Typ::Override.new({content: "#{wpmlns}.fontTable+xml", path: '/word/fontTable.xml'}),
            Typ::Override.new({content: "#{wpmlns}.numbering+xml", path: '/word/numbering.xml'}),
            Typ::Override.new({content: "#{wpmlns}.document.main+xml", path: '/word/document.xml'})
          ]
        })

        # build .rels documents
        relns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        outer = Rel::Relationships.new({rels: [
          Rel::Relationship.new({id: 'rId1', type: "#{relns}/officeDocument", target: 'word/document.xml'})
        ]})
        inner = Rel::Relationships.new({rels: [
          Rel::Relationship.new({id: 'rId1', type: "#{relns}/settings", target: 'settings.xml'}),
          Rel::Relationship.new({id: 'rId2', type: "#{relns}/fontTable", target: 'fontTable.xml'}),
          Rel::Relationship.new({id: 'rId3', type: "#{relns}/numbering", target: 'numbering.xml'}),
          Rel::Relationship.new({id: 'rId4', type: "#{relns}/styles", target: 'styles.xml'})
        ]})

        Dir.mktmpdir do |tmpdir|
          Xml.write_file("#{tmpdir}/[Content_Types].xml", types, xmlns: Typ::Namespace, standalone: true)
          Xml.write_file("#{tmpdir}/.rels", outer, xmlns: Rel::Namespace)
          Xml.write_file("#{tmpdir}/document.xml.rels", inner, xmlns: Rel::Namespace, standalone: true)
          Xml.write_file("#{tmpdir}/document.xml", @document, namespaces: Namespaces)
          Xml.write_file("#{tmpdir}/settings.xml", @settings, standalone: true, namespaces: Namespaces)
          Xml.write_file("#{tmpdir}/fontTable.xml", @font_table, standalone: true, namespaces: Namespaces)
          Xml.write_file("#{tmpdir}/numbering.xml", @numbering, standalone: true, namespaces: Namespaces)
          Xml.write_file("#{tmpdir}/styles.xml", @styles, standalone: true, namespaces: Namespaces)
          Zip::Archive.open(@filename, Zip::CREATE|Zip::TRUNC) do |zip|
            zip.add_file('word/numbering.xml', "#{tmpdir}/numbering.xml")
            zip.add_file('word/settings.xml', "#{tmpdir}/settings.xml")
            zip.add_file('word/fontTable.xml', "#{tmpdir}/fontTable.xml")
            zip.add_file('word/styles.xml', "#{tmpdir}/styles.xml")
            zip.add_file('word/document.xml', "#{tmpdir}/document.xml")
            zip.add_file('word/_rels/document.xml.rels', "#{tmpdir}/document.xml.rels")
            zip.add_file('_rels/.rels', "#{tmpdir}/.rels")
            zip.add_file('[Content_Types].xml', "#{tmpdir}/[Content_Types].xml")
          end
        end
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
        @paragraph = W::Paragraph.new({
          properties: {contextual_spacing: {val: 0}},
          content: [W::Run.new(properties: {right_to_left: {val: 0}})]
        })
      end

      def set_list_style(style_name, indent: 0)
        style = @document.numbering_styles[style_name]
        raise KeyError, "unknown list style '#{style_name}'" if style.nil?
        max_indent = style[:abstract].levels.count - 1
        if indent > max_indent
          raise IndexError, "can't use indent #{indent} with list style '#{style_name}' (max #{max_indent})"
        end

        props = @paragraph.properties
        props.numbering.indent.val = indent
        props.numbering.id.val = style[:definition].id
        # TODO: Is this correct source for this data? Check more example files.
        props.indent.left = style[:abstract].levels[indent].paragraph.indent.left
        props.indent.hanging = Units::Inches * 0.5
        props.run.underline.val = 'none' if indent > 0
        props
      end

      def set_font_size(size)
        run = @paragraph.properties.run
        run.font_size.val = size
        run.font_size_complex.val = size
        run
      end

      def add_text(text)
        run = W::Run.new({
          content: [W::Text.new({space: 'preserve', text: text})],
          properties: {right_to_left: {val: 0}}
        })
        default = @paragraph.properties.run
        run.properties.font_size = default.get_tag(:font_size) # preserves nil
        run.properties.font_size_complex = default.get_tag(:font_size_complex)

        prev = @paragraph.content.last
        @paragraph.content.pop if prev.is_a? W::Run and prev.content.empty?
        @paragraph.content.push(run)
        run
      end
    end
  end
end
