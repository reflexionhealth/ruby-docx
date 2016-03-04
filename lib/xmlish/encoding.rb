module Xmlish
  def self.dump(tag, root: true, standalone: false, schemas: {}, prefixes: nil)
    klass = tag.class
    if prefixes.nil?
      prefixes = {}
      schemas.each { |pre, ns| prefixes[ns] = pre }
    end

    prefix = nil
    unless klass.tag_schema.nil?
      prefix = prefixes[klass.tag_schema]
      raise "prefix for schema '#{klass.tag_schema}' was not provided" if prefix.nil?
      prefix += ':'
    end

    # construct opening/closing tags
    opening = ''
    opening += "<#{prefix}#{klass.tag_name}"
    schemas.each { |pre, ns| opening += " xmlns:#{pre}=\"#{ns}\"" } if root
    klass.tag_attributes.each do |symbol, xmlattr|
      value = tag.send(symbol)
      opening += " #{prefix}#{xmlattr}=\"#{value}\"" unless value.nil?
    end

    # construct inner content
    content = ''
    klass.tag_sequence.each do |symbol|
      child = klass.tag_children[symbol]
      case child[:type]
      when :tag
        value = tag.get_tag(symbol)
        value = child[:class].new if value.nil? and child[:required]
        unless value.nil?
          unless value.is_a? child[:class]
            basename = klass.name.split('::').last || klass.tag_name
            raise TypeError, "expected #{basename}.#{symbol} to be <#{child[:class].tag_name}>::Tag " \
                             "but got #{value.class.name}"
          end
          content += Xmlish.marshal(value, schemas: schemas, root: false)
        end

      when :choice
        value = tag.get_tags(symbol)
        if child[:max] == 1
          if value.nil? and child[:min] > 0
            raise "too few tags provided for '#{symbol}' in #{tag.inspect}:#{klass.name} (0 of 1)"
          end
          content += Xmlish.marshal(value, schemas: schemas, root: false) if value.nil?
        else
          value ||= []
          num, min, max = value.count, child[:min], child[:max]
          if num < min
            range = max > 0 ? "#{min}..#{max}" : "#{min}+"
            raise "too few tags provided for '#{symbol}' in #{tag.inspect}:#{klass.name} (#{num} for #{range})"
          elsif max > 0 and num > max
            raise "too many tags provided for '#{symbol}' in #{tag.inspect}:#{klass.name} (#{num} for #{min}..#{max})"
          end
          value.each { |item| content += Xmlish.marshal(item, schemas: schemas, root: false) }
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
          content += value
        when :tags
          value.each { |item| content += Xmlish.marshal(item, schemas: schemas, root: false) }
        when :mixed
          value.each do |item|
            if item.is_a? Tag
              content += Xmlish.marshal(item, schemas: schemas, root: false)
            else
              content += item.to_s
            end
          end
        end
      end
    end

    closing = ''
    if content.empty?
      opening += '/>'
    else
      opening += '>'
      closing += "</#{prefix}#{klass.tag_name}>"
    end

    # combine output
    output = ''
    output += "<?xml version=\"1.0\" encoding=\"UTF-8\"#{' standalone="yes"' if standalone}?>\n" if root
    output += (opening + content + closing)
    output
  end
end
