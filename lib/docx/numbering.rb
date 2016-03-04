require_relative 'elements'

module Docx
  module Numbering
    include Elements
    include Units

    def table_of_contents
      W::AbstractNumberDefinition.new({
        levels: ((0...9).map do |level|
          W::LevelDefinition.new({
            level: level,
            start: {val: 1},
            format: {val: 'decimal'},
            text: {val: (0...level).map { |x| "%#{x + 1}." }.join},
            justify: {val: 'right'},
            paragraph: {indent: {left: halfinch * (level + 1), first_line: halfinch * (level + 0.5)}},
            run: {underline: 'none'}
          })
        end)
      })
    end
  end
end
