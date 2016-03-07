require_relative 'size'
require_relative 'constants'
require_relative 'elements'

module Docx
  module Styles
    include Elements

    def self.default
      W::DefaultProperties.new({
        run: {properties: {
          fonts: {ascii: 'Arial', complex: 'Arial', east_asia: 'Arial', ansi: 'Arial'},
          bold: {val: Docx::False},
          italic: {val: Docx::False},
          small_caps: {val: Docx::False},
          strike: {val: Docx::False},
          color: {val: Docx::Black},
          font_size: {val: Points * 11},
          font_size_complex: {val: Points * 11},
          underline: {val: 'none'},
          vertical_align: {val: 'baseline'}
        }},
        paragraph: {properties: {
          keep_next: {val: Docx::False},
          keep_lines: {val: Docx::False},
          widow_control: {val: Docx::True},
          spacing: {after: 0, before: 0, line: Halfpts * 276, line_rule: 'auto'},
          indent: {left: 0, right: 0, first_line: 0},
          justify: {val: 'left'}
        }}
      })
    end

    def self.paragraph
      W::Style.new({
        type: 'paragraph',
        id: 'Normal',
        default: 1,
        name: {val: 'normal'}
      })
    end

    def self.table
      W::Style.new({
        type: 'table',
        id: 'TableNormal',
        default: 1,
        name: {val: 'Table Normal'}
      })
    end

    def self.h1
      W::Style.new({
        type: 'paragraph',
        id: 'Heading1',
        name: {val: 'heading 1'},
        based_on: {val: 'Normal'},
        next: {val: 'Normal'},
        paragraph: {
          keep_next: {val: Docx::True},
          keep_lines: {val: Docx::True},
          spacing: {after: Halfpts * 120, before: Halfpts * 400, line_rule: 'auto'},
          contextual_spacing: {val: Docx::True}
        },
        run: {
          font_size: {val: Points * 20},
          font_size_complex: {val: Points * 20}
        }
      })
    end

    def self.h2
      W::Style.new({
        type: 'paragraph',
        id: 'Heading2',
        name: {val: 'heading 2'},
        based_on: {val: 'Normal'},
        next: {val: 'Normal'},
        paragraph: {
          keep_next: {val: Docx::True},
          keep_lines: {val: Docx::True},
          spacing: {after: Halfpts * 120, before: Halfpts * 360, line_rule: 'auto'},
          contextual_spacing: {val: Docx::True}
        },
        run: {
          bold: {val: Docx::False},
          font_size: {val: Points * 16},
          font_size_complex: {val: Points * 16}
        }
      })
    end

    def self.h3
      W::Style.new({
        type: 'paragraph',
        id: 'Heading3',
        name: {val: 'heading 3'},
        based_on: {val: 'Normal'},
        next: {val: 'Normal'},
        paragraph: {
          keep_next: {val: Docx::True},
          keep_lines: {val: Docx::True},
          spacing: {after: Halfpts * 80, before: Halfpts * 320, line_rule: 'auto'},
          contextual_spacing: {val: Docx::True}
        },
        run: {
          bold: {val: Docx::False},
          color: {val: '434343'},
          font_size: {val: Points * 14},
          font_size_complex: {val: Points * 14}
        }
      })
    end

    def self.h4
      W::Style.new({
        type: 'paragraph',
        id: 'Heading4',
        name: {val: 'heading 4'},
        based_on: {val: 'Normal'},
        next: {val: 'Normal'},
        paragraph: {
          keep_next: {val: Docx::True},
          keep_lines: {val: Docx::True},
          spacing: {after: Halfpts * 80, before: Halfpts * 280, line_rule: 'auto'},
          contextual_spacing: {val: Docx::True}
        },
        run: {
          color: {val: Docx::Gray},
          font_size: {val: Points * 12},
          font_size_complex: {val: Points * 12}
        }
      })
    end

    def self.h5
      W::Style.new({
        type: 'paragraph',
        id: 'Heading5',
        name: {val: 'heading 5'},
        based_on: {val: 'Normal'},
        next: {val: 'Normal'},
        paragraph: {
          keep_next: {val: Docx::True},
          keep_lines: {val: Docx::True},
          spacing: {after: Halfpts * 80, before: Halfpts * 240, line_rule: 'auto'},
          contextual_spacing: {val: Docx::True}
        },
        run: {
          color: {val: Docx::Gray},
          font_size: {val: Points * 11},
          font_size_complex: {val: Points * 11}
        }
      })
    end

    def self.h6
      W::Style.new({
        type: 'paragraph',
        id: 'Heading6',
        name: {val: 'heading 6'},
        based_on: {val: 'Normal'},
        next: {val: 'Normal'},
        paragraph: {
          keep_next: {val: Docx::True},
          keep_lines: {val: Docx::True},
          spacing: {after: Halfpts * 80, before: Halfpts * 240, line_rule: 'auto'},
          contextual_spacing: {val: Docx::True}
        },
        run: {
          italic: {val: Docx::True},
          color: {val: Docx::Gray},
          font_size: {val: Points * 11},
          font_size_complex: {val: Points * 11}
        }
      })
    end

    def self.title
      W::Style.new({
        type: 'paragraph',
        id: 'Title',
        name: {val: 'Title'},
        based_on: {val: 'Normal'},
        next: {val: 'Normal'},
        paragraph: {
          keep_next: {val: Docx::True},
          keep_lines: {val: Docx::True},
          spacing: {after: Halfpts * 60, before: 0, line_rule: 'auto'},
          contextual_spacing: {val: Docx::True}
        },
        run: {
          font_size: {val: Points * 26},
          font_size_complex: {val: Points * 26}
        }
      })
    end

    def self.subtitle
      W::Style.new({
        type: 'paragraph',
        id: 'Subtitle',
        name: {val: 'Subtitle'},
        based_on: {val: 'Normal'},
        next: {val: 'Normal'},
        paragraph: {
          keep_next: {val: Docx::True},
          keep_lines: {val: Docx::True},
          spacing: {after: Halfpts * 320, before: 0, line_rule: 'auto'},
          contextual_spacing: {val: Docx::True}
        },
        run: {
          fonts: {ascii: 'Arial', complex: 'Arial', east_asia: 'Arial', ansi: 'Arial'},
          italic: {val: Docx::False},
          color: {val: Docx::Gray},
          font_size: {val: Points * 15},
          font_size_complex: {val: Points * 15}
        }
      })
    end
  end
end
