module Acrobat

  class PDoc

    attr_reader :app, :ole_obj, :path

    def initialize(app,ole,path=nil)
      @app = app
      @ole_obj = ole
      @path = path
    end

    def show(name = nil)
      name = name ||  ole_obj.GetFileName
      ole_obj.OpenAVDoc(name)
    end

    # @return [Fixnum] the number of pages in the pdf
    def page_count
      ole_obj.GetNumPages()
    end

    # merges the doc to the
    # @overload merge(doc)
    #   @param doc [String] the String path of a pdf file
    # @overload merge(doc)
    #   @param doc [PDoc] an open PDoc to merge
    # @return [Boolean] whether the doc was merged correctly
    def merge(doc)
      case doc
      when PDoc
        merge_pdoc(doc)
      when String, Pathname
        docpath = Pathname(doc)
        raise 'File not found' unless docpath.file?
        doc2 = app.open(docpath)
        merge_pdoc(doc2)
      end
    end

    def fill_and_save(results,name: nil, dir: nil)
      fill_form(results)
      is_saved = save_as(name: name, dir: dir)
      puts "saved file: %s" % [dir + name] if is_saved
      true
    end

    def default_dir(d)
      Pathname(dir || Pathname.getw)
    end

    def save_as(name:nil, dir:nil)
      name = path.basename unless name
      dir = Pathname(dir || Pathname.getwd)
      dir.mkpath
      windows_path = FileSystemObject.windows_path(dir + name )
      ole_obj.save(ACRO::PDSaveFull | ACRO::PDSaveCopy,windows_path)
    end

    def name
      ole_obj.GetFileName
    end

    def close
      ole_obj.Close rescue nil
    end


    def jso
      @jso ||= Jso.new(self,ole_obj.GetJSObject)
    end

    def field_names
      jso.field_names
    end

    def fill_form(results)
      jso.fill_form(results)
    end

    protected
    def merge_pdoc(doc)
      begin
        merged = ole_obj.InsertPages(page_count - 1, doc.ole_obj, 0, doc.page_count, true)
        return merged
      rescue
        return false
      end
    end
  end

end
