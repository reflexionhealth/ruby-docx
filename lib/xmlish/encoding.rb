module Xmlish
  def self.dump(tag, root = false, standalone: false, xmlns: nil, namespaces: {})
    buffer = StringIO.new
    encoder = Encoder.new(buffer, xmlns: xmlns, namespaces: namespaces)
    encoder.encode(tag, root, standalone: standalone)
    buffer.string
  end

  def self.write_file(path, tag, standalone: false, xmlns: nil, namespaces: {})
    written = 0
    File.open(path, 'w') do |file|
      encoder = Encoder.new(file, xmlns: xmlns, namespaces: namespaces)
      written = encoder.encode(tag, true, standalone: standalone)
    end
    written
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
      @written = 0
      if root
        directives = 'version="1.0" encoding="UTF-8"'
        directives += ' standalone="yes"' if standalone
        write("<?xml #{directives}?>\r")
        recurse(tag, true)
      else
        recurse(tag, false)
      end
      @written
    end

  private
    def write(string)
      @written += @file.write(string)
    end

    def recurse(tag, xmlns_attrs = false)
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

      # write opening tag
      write("<#{prefix}#{klass.tag_type}")
      if xmlns_attrs
        write(" xmlns=\"#{@xmlns}\"") unless @xmlns.nil?
        @namespaces.each { |pre, ns| write(" xmlns:#{pre}=\"#{ns}\"") }
      end
      klass.tag_attributes.each do |symbol, info|
        value = tag.send(symbol)
        attr_prefix = info[:prefix] ? "#{info[:prefix]}:" : prefix
        write(" #{attr_prefix}#{info[:name]}=\"#{value}\"") unless value.nil?
      end

      # collect inner tags/content
      content = []
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
            content.push(value)
          end

        when :choice
          value = tag.get_tags(symbol)
          if child[:max] == 1
            if !value.nil?
              content.push(value)
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
            value.each { |item| content.push(item) }
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
            content.push(value)
          when :tags, :mixed
            value.each { |item| content.push(item) }
          end
        end
      end

      if content.empty?
        write('/>')
      else
        write('>')
        # write inner xml
        content.each { |item| item.is_a?(Tag) ? recurse(item) : write(item.to_s) }
        # write closing tag
        write("</#{prefix}#{klass.tag_type}>")
      end
    end
  end
end
