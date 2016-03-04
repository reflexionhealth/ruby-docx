module Docx
  class Tag
    ContentTypes = [:tags, :text, :mixed]

    # Specifies the XML name of this tag (eg. <title>).
    def self.title(title)
      self.tag_name = title
    end

    # Specifies the XML schema of this tag (eg. http://schemas.openxmlformats.org/drawingml/2006/main)
    def self.schema(schema)
      self.tag_schema = schema
    end

    # Specifies new XML attributes for this tag.
    #
    #     attributes :color, :depth
    #
    # Each attribute's xml name will match the +to_s+ of the symbol.
    # To define an attribute with a custom xml name, use +attribute+ instead.
    def self.attributes(*symbols)
      attrs = symbols.each_with_object({}) { |sym, hash| hash[sym] = sym.to_s.split('_').map(&:capitalize).join }
      self.tag_attributes.merge! attrs
      symbols.each { |sym| attr_accessor sym }
    end

    # Specifies a new XML attribute for this tag with an optional xml name
    #
    #     attribute :color
    #     attribute :revision_id, 'rsid'
    #
    # See also: +attributes+
    def self.attribute(symbol, xmlattr = nil)
      self.tag_attributes[symbol] = xmlattr || symbol.to_s.split('_').map(&:capitalize).join
      attr_accessor symbol
    end

    # Specifies a block in the nested tag content that can be arbitrary text, tags, or a mixture.
    # An attr_accessor is provided using the first argument as an attribute name.
    #
    #     content :attrname, :tags
    #
    # The options for calling content are :tags, :text, or :mixed
    def self.content(symbol, type)
      unless [:text, :mixed, :tags].include?(type)
        raise "unknown content type '#{type}', must be :text, :tags, or :mixed"
      end
      variable = "@#{symbol}".to_sym
      self.tag_children[symbol] = {type: :content, variable: variable, content: type}
      self.tag_sequence.push(symbol)

      attr_writer symbol
      define_method(symbol) do
        value = self.instance_variable_get(variable)
        if value.nil?
          self.instance_variable_set(variable, type == :text ? '' : [])
        else
          value
        end
      end
    end

    # Adds a child tag to the end of this tag's sequence.
    #
    #     tag :properties, SectionProperties
    #     tag :header, TagHeader, required: true
    #
    # See also: +tags+
    def self.tag(symbol, klass, required: false)
      variable = "@#{symbol}".to_sym
      self.tag_children[symbol] = {type: :tag, variable: variable, class: klass, required: required}
      self.tag_sequence.push(symbol)

      attr_writer symbol
      define_method(symbol) do
        value = self.instance_variable_get(variable)
        value ? value : self.instance_variable_set(variable, klass.new)
      end
    end

    # Adds a choice of tags to the end of this tag's sequence.
    # Allows a tag to contain an unbounded list of child tags.
    #
    #    tags :children, Child
    #    tags :plants, [Tree, Grass, Bush]
    #
    # If +max+ is negative (default), there is no limit to the total number of tags
    # If +max+ is exactly 1, the items are a list of mutually exclusive tags
    #
    #    tags :exclusive_required, [A, B, C], min: 1, max: 1
    #    tags :exclusive_optional, [A, B, C], min: 0, max: 1
    #    tags :some_things, Thing, min: 1
    #    tags :upto_six_things, Thing, max: 6
    #
    # See also +tag+
    def self.tags(symbol, classes, min: 0, max: -1)
      if classes.is_a? Class
        return self.tag(symbol, classes, required: min > 0) if max == 1
        classes = [classes]
      end
      variable = "@#{symbol}".to_sym
      self.tag_children[symbol] = {type: :choice, variable: variable, choices: classes, min: min, max: max}
      self.tag_sequence.push(symbol)

      attr_writer symbol
      define_method(symbol) do
        value = self.instance_variable_get(variable)
        if value.nil?
          self.instance_variable_set(variable, (max == 1) ? klass.new : [])
        else
          value
        end
      end
    end

    def self.inherited(subclass)
      subclass.singleton_class.instance_eval do
        attr_accessor :tag_name
        attr_accessor :tag_schema
        attr_accessor :tag_attributes
        attr_accessor :tag_children
        attr_accessor :tag_sequence
      end
      subclass.tag_name = subclass.name[0].downcase + subclass.name[1..-1] if subclass.name
      subclass.tag_attributes = {}
      subclass.tag_children = {}
      subclass.tag_sequence = []
    end

    def initialize(attrs = {})
      attrs.each do |name, value|
        if self.class.tag_attributes.key? name
          self.send("#{name}=", value)
        elsif self.class.tag_children[name]
          child = self.class.tag_children[name]
          case child[:type]
          when :tag
            if value.is_a? Hash
              self.send("#{name}=", child[:class].new(value))
            else
              self.send("#{name}=", value)
            end
          else
            self.send("#{name}=", value)
          end
        else
          raise "unknown attribute or tag '#{name}' for #{self.class.name}"
        end
      end
    end

    def get_tag(sym)
      child = self.tag_children[sym]
      raise "undefined tag '#{sym}' for #{self.class.name}" if child.nil?
      self.instance_variable_get(child[:variable])
    end
    alias_method :get_tags, :get_tags

    def set_tag(sym, value)
      child = self.tag_children[sym]
      raise "undefined tag '#{sym}' for #{self.class.name}" if child.nil?
      self.instance_variable_set(child[:variable])
    end
    alias_method :set_tags, :set_tag

    def inspect
      pretty = "<#{self.class.name}"
      self.class.tag_attributes.each do |symbol, _|
        value = self.send(symbol)
        pretty += " #{symbol}=\"#{value}\"" unless value.nil?
      end
      pretty += '>'
      pretty
    end
  end
end
