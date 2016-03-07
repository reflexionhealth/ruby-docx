module Docx
  # Provides complex representation of a Size.
  # It can either be used with extensions to numeric:
  #
  #     Numeric.extend(Docx::NumericExt)
  #
  #     8.pt
  #     11.inches
  #
  # Or used with Size constants:
  #
  #     Docx::Inches * 10
  #
  class Size
    attr_reader :value, :parts

    # e.g. Size.new(10 * Docx::Inches, [[:inches, 10]])
    def initialize(value, parts)
      @value, @parts = value, parts
    end

    # Adds another Size or a Numeric to this Size.
    # Numeric values are treated as halfpts.
    def +(other)
      if Size === other
        Size.new(value + other.value, @parts + other.parts)
      else
        Size.new(value + other, @parts + [[:halfpts, other]])
      end
    end

    def +@
      self
    end

    # Subtracts another Size or a Numeric from this Size.
    # Numeric values are treated as halfpts.
    def -(other)
      self + (-other)
    end

    def -@
      Size.new(-value, @parts.map { |unit, amount| [unit, -amount] })
    end

    # Multiplies this size by a Numeric.
    def *(other)
      scaled_parts = @parts.map do |unit, amount|
        scaled = amount * other
        scaled = scaled.to_i if scaled.is_a? Float and scaled == scaled.to_i.to_f
        [unit, scaled]
      end
      Size.new((value * other).round, scaled_parts)
    end

    def /(other)
      scaled_parts = @parts.map do |unit, amount|
        scaled = amount / other
        scaled = scaled.to_i if scaled.is_a? Float and scaled == scaled.to_i.to_f
        [unit, scaled]
      end
      Size.new((value / other).round, scaled_parts)
    end

    # Returns +true+ if +other+ is also a Size instance, which has the same parts as this one.
    def eql?(other)
      Size === other && other.value.eql?(value)
    end

    def self.===(other)
      other.is_a?(Size)
    rescue ::NoMethodError
      false
    end

    def to_i; @value.to_i; end
    def to_f; @value.to_f; end
    def to_s; @value.to_s; end
    def hash; @value.hash; end
    def is_a?(klass); Size == klass || @value.is_a?(klass); end
    def instance_of?(klass); Size == klass || @value.instance_of?(klass); end

    def inspect
      singular = {inches: 'inch'}
      amount_per_unit = @parts.each_with_object(Hash.new(0)) { |(unit, amount), hash| hash[unit] += amount }
      texts = amount_per_unit.sort_by { |unit, _| [:inches, :points, :halfpts].index(unit) }
                             .map { |unit, amount| "#{amount} #{amount == 1 ? singular[unit] || unit[0...-1] : unit}" }
      texts[0...-1].each_with_index { |text, i| texts[i] = text + ',' } if texts.count > 2
      texts[-1] = 'and ' + texts[-1] if texts.count > 1
      texts.join(' ')
    end
  end

  module NumericExt
    def self.extended(numeric)
      numeric.class_eval do
        def pt; Docx::Size.new(self * Docx::Point, [[:points, self]]); end
        def inches; Docx::Size.new(self * Docx::Inch, [[:inches, self]]); end
      end
    end
  end
end
