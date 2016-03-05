require_relative 'elements'

module Docx
  module Numbering
    include Elements
    include Units

    def self.table_of_contents
      halfinch = Inches * 0.5
      W::AbstractNumberDefinition.new({
        levels: ((0...9).map do |level|
          W::LevelDefinition.new({
            level: level,
            start: {val: 1},
            format: {val: 'decimal'},
            text: {val: (1..level + 1).map { |x| "%#{x}." }.join},
            justify: {val: 'right'},
            paragraph: {indent: {left: halfinch * (level + 1), first_line: halfinch * (level + 0.5)}},
            run: {underline: {val: 'none'}}
          })
        end)
      })
    end
  end
end
