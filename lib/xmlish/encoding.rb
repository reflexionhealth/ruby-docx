module Xmlish
  def self.dump(tag, root = false, standalone: false, xmlns: nil, namespaces: {})
    file = StringIO.new
    encoder = Encoder.new(file, xmlns: xmlns, namespaces: namespaces)
    encoder.encode(tag, root, standalone: standalone)
    file.string
  end

  # TODO: Move TypeErrors into the attr_writer?
  class EncodeError < StandardError; end
  class Encoder
    def initialize(file, xmlns: nil, namespaces: {})
      @file = file
      @xmlns = xmlns
      @namespaces = namespaces
      @prefixes = namespaces.each_with_object({}) { |(pre, ns), hash| hash[ns] = pre }
    end

    def encode(tag, root = false, standalone: false)
      klass = tag.class
      prefix = nil
      namespace = klass.tag_namespace
      unless namespace.nil?
        if namespace != @xmlns
          prefix = @prefixes[namespace]
          raise EncodeError, "prefix for non-default namespace '#{namespace}' was not provided" if prefix.nil?
          prefix += ':'
        end
      end

      # write xml directive
      if root
        directives = 'version="1.0" encoding="UTF-8"'
        directives += ' standalone="yes"' if standalone
        @file.write("<?xml #{directives}?>\r")
      end

      # write opening tag
      @file.write("<#{prefix}#{klass.tag_type}")
      @namespaces.each { |pre, ns| @file.write(" xmlns:#{pre}=\"#{ns}\"") } if root
      klass.tag_attributes.each do |symbol, xmlattr|
        value = tag.send(symbol)
        @file.write(" #{prefix}#{xmlattr}=\"#{value}\"") unless value.nil?
      end

      empty = klass.tag_sequence.empty?
      empty ? @file.write('/>') : @file.write('>')

      # write inner xml
      content = ''
      klass.tag_sequence.each do |symbol|
        child = klass.tag_children[symbol]
        case child[:type]
        when :tag
          value = tag.get_tag(symbol)
          value = child[:class].new if value.nil? and child[:required]
          unless value.nil?
            unless value.is_a? child[:class]
              basename = klass.name.split('::').last || klass.tag_type
              raise TypeError, "expected #{basename}.#{symbol} to be <#{child[:class].tag_type}>::Tag " \
                               "but got #{value.class.name}"
            end
            self.encode(value)
          end

        when :choice
          value = tag.get_tags(symbol)
          if child[:max] == 1
            if !value.nil?
              self.encode(value)
            elsif child[:min] > 0
              raise EncodeError, "too few tags provided for '#{symbol}' in #{tag.inspect}:#{klass.name} (0 of 1)"
            end
          else
            value ||= []
            num, min, max = value.count, child[:min], child[:max]
            if num < min
              range = max > 0 ? "#{min}..#{max}" : "#{min}+"
              raise EncodeError, "too few tags provided for '#{symbol}' in #{tag.inspect}:#{klass.name} (#{num} for #{range})"
            elsif max > 0 and num > max
              raise EncodeError, "too many tags provided for '#{symbol}' in #{tag.inspect}:#{klass.name} (#{num} for #{min}..#{max})"
            end
            value.each { |item| self.encode(item) }
          end

        when :content
          value = tag.get_content(symbol)
          next if value.nil?
          case child[:content]
          when :text
            unless value.is_a? String
              basename = klass.name.split('::').last
              raise TypeError, "expected #{basename}.#{symbol} to be String but got #{value.class.name}"
            end
            @file.write(value)
          when :tags
            value.each { |item| self.encode(item) }
          when :mixed
            value.each { |item| item.is_a?(Tag) ? self.encode(item) : @file.write(item.to_s) }
          end
        end
      end

      # write closing tag
      unless empty
        @file.write("</#{prefix}#{klass.tag_type}>")
      end
    end
  end
end
