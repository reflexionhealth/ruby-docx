require_relative 'size'

module Docx
  # Bools
  True = 1
  False = 0

  # Colors
  Black = '000000'.freeze
  Gray  = '666666'.freeze # spelling arguments ensue
  White = 'FFFFFF'.freeze

  # Units
  Halfpt = 1
  Point = 2
  Inch = 1440

  # Sizes
  Halfpts = Size.new(Halfpt, [[:halfpts, 1]])
  Points = Size.new(Point, [[:points, 1]])
  Inches = Size.new(Inch, [[:inches, 1]])
end
