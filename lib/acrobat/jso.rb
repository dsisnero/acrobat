module Acrobat
  class Jso
    attr_reader :doc, :ole_obj

    def initialize(doc, ole)
      @doc = doc
      @ole_obj = ole
    end

    def find_field(name_or_number)
      case name_or_number
      when String, Symbol
        ole_get_field(name_or_number.to_s)
      when Number
        ole_get_field(name_or_number)
      end
    end

    def ole_get_field(field)
      ole_obj.getField(field)
    end

    def console
      @console ||= ole_obj.console
    end

    def show_console
      console.show
    end

    def field_names
      result = []
      count = field_count
      0.upto(count - 1) do |i|
        result << ole_obj.getNthFieldName(i)
      end
      result
    end

    def export_as_fdf(name)
    end

    def import_fdf(path)
    end

    def fields_hash
      result = {}
      field_names.each_with_object(result) do |name, h|
        h[name] = get_field(name)
      end
    end


    # // Enumerate through all of the fields in the document.
    # for (var i = 0; i < this.numFields; i++)
    # console.println("Field[" + i + "] = " + this.getNthFieldName(i));

    def set_field(name, value)
      field = find_field(name)
      field.Value = value.to_s if field
    rescue
      require "pry"
      binding.pry
      nil
    end

    def get_field(name)
      field = find_field(name)
      field.Value if field
    end

    def field_count
      ole_obj.numFields.to_int
    end

    def clear_form
      ole_obj.resetForm
    end

    def fill_form(hash)
      clear_form
      hash.each do |k, v|
        set_field(k, v)
      end
    end
  end
end
