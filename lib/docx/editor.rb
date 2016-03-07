require_relative 'bool'
require_relative 'units'
require_relative 'elements'
require_relative 'numbering'
require_relative 'styles'
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
      doc.document.background.color = Color::White
      doc.set_page_size(width: Inches * 8.5, height: Inches * 11)
      doc.set_page_margins(top: Inches * 1, bottom: Inches * 1, left: Inches * 1, right: Inches * 1)
      doc.set_page_numbering(start: 1)
      doc.font_table.fonts.push(Elements::W::Font.new(name: 'Arial'))
      doc.settings.display_background_shapes.val = Docx::True
      doc.settings.default_tab_stop.val = Inches * 0.5
      doc.styles.default = Styles.default
      doc.define_font_style(:default, Styles.paragraph)
      doc.define_table_style(:default, Styles.table)
      doc.define_font_style(:h1, Styles.h1)
      doc.define_font_style(:h2, Styles.h2)
      doc.define_font_style(:h3, Styles.h3)
      doc.define_font_style(:h4, Styles.h4)
      doc.define_font_style(:h5, Styles.h5)
      doc.define_font_style(:h6, Styles.h6)
      doc.define_font_style(:title, Styles.title)
      doc.define_font_style(:subtitle, Styles.subtitle)
      doc
    end

    class Document
      include Elements

      attr_accessor :filename
      attr_accessor :document
      attr_accessor :font_table
      attr_accessor :numbering
      attr_accessor :settings
      attr_accessor :styles

      attr_reader :bookmarks
      attr_reader :font_styles
      attr_reader :table_styles
      attr_reader :numbering_styles

      def initialize
        @document = W::Document.new
        @font_table = W::FontTable.new
        @numbering = W::Numbering.new

        compat = {val: 14, name: 'compatibilityMode', uri: 'http://schemas.microsoft.com/office/word'}
        @settings = W::Settings.new(compatibility: {setting: compat})
        @styles = W::Styles.new

        @bookmarks = []
        @font_styles = {}
        @list_styles = {}
        @table_style = {}
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
          Xml.write_file("#{tmpdir}/.rels", outer, xmlns: Rel::Namespace, standalone: true)
          Xml.write_file("#{tmpdir}/document.xml.rels", inner, xmlns: Rel::Namespace, standalone: true)
          Xml.write_file("#{tmpdir}/document.xml", @document, namespaces: Namespaces)
          Xml.write_file("#{tmpdir}/settings.xml", @settings, standalone: true, namespaces: Namespaces)
          Xml.write_file("#{tmpdir}/fontTable.xml", @font_table, standalone: true, namespaces: Namespaces)
          Xml.write_file("#{tmpdir}/numbering.xml", @numbering, standalone: true, namespaces: Namespaces)
          Xml.write_file("#{tmpdir}/styles.xml", @styles, standalone: true, namespaces: Namespaces)
          Zip::Archive.open(@filename, Zip::CREATE | Zip::TRUNC) do |zip|
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

      def define_font_style(style_name, font_style)
        doc.styles.styles.push(font_style)
        @font_styles[style_name] = font_style
      end

      def define_table_style(style_name, table_style)
        doc.styles.styles.push(table_style)
        @table_styles[style_name] = table_style
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
        @document.body.content.push(paragraph.root)
        paragraph
      end

      def add_page_break
        pg = W::Paragraph.new(content: [W::Run.new(content: [W::Break.new(type: 'page')])])
        @document.body.content.push(pg)
        pg
      end

      def add_horizontal_rule
        pg = W::Paragraph.new(properties: {border: {top: {color: 'auto', space: 1, sz: 4, val: 'single'}}})
        @document.body.content.push(pg)
        pg
      end

      def add_table
        table = Table.new(self)
        @document.body.content.push(table.root)
        table
      end
    end

    class Paragraph
      include Elements

      attr_accessor :document
      attr_accessor :root

      def initialize(editor_doc)
        @document = editor_doc
        @root = W::Paragraph.new({
          properties: {contextual_spacing: {val: Docx::False}},
          content: [W::Run.new(properties: {right_to_left: {val: Docx::False}})]
        })
      end

      def set_list_style(style_name, indent: 0)
        style = @document.numbering_styles[style_name]
        raise KeyError, "unknown list style '#{style_name}'" if style.nil?
        max_indent = style[:abstract].levels.count - 1
        if indent > max_indent
          raise IndexError, "can't use indent #{indent} with list style '#{style_name}' (max #{max_indent})"
        end

        props = @root.properties
        props.contextual_spacing.val = Docx::True
        props.numbering.indent.val = indent
        props.numbering.id.val = style[:definition].id
        # TODO: Is this correct source for this data? Check more example files.
        props.indent.left = style[:abstract].levels[indent].paragraph.indent.left
        props.indent.hanging = Units::Inches / 4
        props.run.underline.val = 'none' if indent > 0
        props
      end

      def set_font_size(size)
        run = @root.properties.run
        run.font_size.val = size
        run.font_size_complex.val = size
        run
      end

      def set_indent(left: nil, right: nil, hanging: nil, first_line: nil)
        indent = @root.properties.indent
        indent.left = left if left
        indent.right = right if right
        indent.hanging = hanging if hanging
        indent.first_line = first_line if first_line
        indent
      end

      def add_bookmark(name)
        id = @document.bookmarks.count
        markstart = W::BookmarkStart.new(col_first: 0, col_last: 0, name: 0, id: id)
        markend = W::BookmarkEnd.new(id: id)
        @root.content.push(markstart)
        @root.content.push(markend)
        @document.bookmarks.push(markstart)
        markstart
      end

      def add_text(text)
        run = self.begin_run
        self.write_text(text)
        run
      end

      def begin_run
        default = @root.properties.run
        run = W::Run.new({properties: {right_to_left: {val: Docx::False}}})
        run.properties.font_size = default.get_tag(:font_size) # preserves nil
        run.properties.font_size_complex = default.get_tag(:font_size_complex)

        prev = @root.content.last
        @root.content.pop if prev.is_a?(W::Run) and prev.content.empty?
        @root.content.push(run)
        run
      end

      def write_text(text)
        run = @root.content.last
        raise "call 'begin_run' prior to 'write_text'" unless run.is_a?(W::Run)
        run.content.push(W::Text.new(space: 'preserve', text: text))
        run
      end

      def write_tab
        run = @root.content.last
        raise "call 'begin_run' prior to 'write_tab'" unless run.is_a?(W::Run)
        run.content.push(W::Tab.new)
        run
      end
    end

    class Table
      include Elements

      attr_accessor :document
      attr_accessor :root

      def initialize(editor_doc)
        @document = editor_doc
        @root = W::Table.new({
          properties: {
            right_to_left: {val: Docx::False},
            justify: {val: 'left'},
            borders: {
              top: {color: Color::Black, space: 0, sz: 8, val: 'single'},
              left: {color: Color::Black, space: 0, sz: 8, val: 'single'},
              bottom: {color: Color::Black, space: 0, sz: 8, val: 'single'},
              right: {color: Color::Black, space: 0, sz: 8, val: 'single'},
              horizontal: {color: Color::Black, space: 0, sz: 8, val: 'single'},
              vertical: {color: Color::Black, space: 0, sz: 8, val: 'single'}
            },
            layout: {type: 'fixed'},
            look: {val: '0600'}
          }
        })
      end

      def set_width(size)
        # NOTE(Kevin): Supposedly "dxa" is 20ths of a point, but Google Docs
        # treats it as a half point, so we'll just ignore all that
        props = @root.properties
        props.width.type = 'dxa'
        props.width.w = size.to_f
        props
      end

      def set_indent(size)
        # NOTE(Kevin): Supposedly "dxa" is 20ths of a point, but Google Docs
        # treats it as a half point, so we'll just ignore all that
        props = @root.properties
        props.indent.type = 'dxa'
        props.indent.w = size.to_f
        props
      end

      def define_columns(widths)
        grid = @root.grid
        grid.columns = widths.map { |w| W::GridColumn.new(width: w) }
        grid.change.previous.columns = grid.columns
        grid
      end

      def add_row
        row = TableRow.new(self)
        @root.content.push(row.root)
        row
      end
    end

    class TableRow
      include Elements

      attr_accessor :table
      attr_accessor :root

      def initialize(editor_tbl)
        @table = editor_tbl
        @root = W::TableRow.new
      end

      def add_cell
        cell = TableCell.new(@table)
        @root.content.push(cell.root)
        cell
      end
    end

    class TableCell
      include Elements

      attr_accessor :table
      attr_accessor :root

      def initialize(editor_tbl)
        @table = editor_tbl
        @root = W::TableCell.new
        self.set_margins(top: 100, right: 100, bottom: 100, left: 100) # ~0.69 inches
      end

      def set_margins(top: nil, right: nil, bottom: nil, left: nil)
        # NOTE(Kevin): Supposedly "dxa" is 20ths of a point, but Google Docs
        # treats it as a half point, so we'll just ignore all that
        margins = @root.properties.margins
        if top
          margins.top.type = 'dxa'
          margins.top.w = top.to_f
        end
        if right
          margins.right.type = 'dxa'
          margins.right.w = right.to_f
        end
        if bottom
          margins.bottom.type = 'dxa'
          margins.bottom.w = bottom.to_f
        end
        if left
          margins.left.type = 'dxa'
          margins.left.w = left.to_f
        end
        margins
      end

      def add_paragraph
        paragraph = Paragraph.new(self)
        @root.content.push(paragraph.root)
        paragraph
      end

      def add_table
        table = Table.new(self)
        @root.content.push(table.root)
        table
      end
    end
  end
end
