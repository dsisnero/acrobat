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

    def last_page
      page_count -1
    end

    # merges the doc to the
    # @overload merge(doc)
    #   @param doc [String] the String path of a pdf file
    # @overload merge(doc)
    #   @param doc [PDoc] an open PDoc to merge
    # @return [Boolean] whether the doc was merged correctly
    def merge(doc)
      src = open_pdoc(doc)
      merge_pdoc(src)
    end

    # opens and/or returns PDoc
    # @overload open(doc)
    #   @param doc [String] the String path of a pdf file
    # @overload open(PDoc] and open PDoc file
    #   @param doc [PDoc] an already open PDoc file
    # @return doc [PDOC] the opened PDoc
    def open_pdoc(doc)
      src = case doc
            when PDoc
              doc
            when String, Pathname
              docpath = Pathname(doc)
              raise 'File not found' unless docpath.file?
              doc2 = app.open(docpath)
              doc2
            end
      src
    end

    def fill_and_save(results,name: nil, dir: nil)
      fill_form(results)
      is_saved = save_as(name: name, dir: dir)
      puts "saved file: %s" % [dir + name] if is_saved
      true
    end

    def insert_pages(src: , insert_after: nil, src_page_start: nil, src_page_end: nil)
      insert_hash = { 'nPage' => insert_after -1 }
      insert_hash['nStart'] = src_page_start + 1 if src_page_start
      insert_hash['nEnd'] = src_page_end + 1 if src_page_end
      ole_obj.InsertPages(**insert_hash)
    end

    def prepend_pages(src_path: , src_page_start: 1, src_page_end: nil)
      insert_pages( insert_after: 0, src_path: src_path, src_page_start: src_page_start, src_page_end: src_page_end)
    end

    # returns [Pathname] of d
    #   @param dir [String, Nil] the String path
    # @return dir [Pathname] Pathname of dir or of working directory
    def default_dir(dir)
      Pathname(dir || Pathname.getw)
    end

    def save_as(name,dir:nil)
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

    def replace_pages(doc, start: 0, replace_start: 0, num_of_pages: 1,merge_annotations: true)
      src = open_pdoc(doc)

      ole_obj.ReplacePages(start,src.ole_obj,replace_start,num_of_pages,merge_annotations)
    end

    # return the instance of JSO object
    # @return [Jso] a WIN32OLE wrapped Jso object 'javascript object'
    def jso
      @jso ||= Jso.new(self,ole_obj.GetJSObject)
    end

    # return the field_names of the pdf form
    def field_names
      jso.field_names
    end

    def fill_form(results)
      jso.fill_form(results)
    end

    protected
    def merge_pdoc(doc, **options)
      begin
        unless options
          merged = ole_obj.InsertPages(page_count - 1, doc.ole_obj, 0, doc.page_count, true)
          return merged
        else
          start = options[:start]
          start_page
          pages = options[:pages]
          ole_obj.InsertPages(start, doc.ole_obj, 0, doc.page_count, true)
        end
      rescue
        return false
      end
    end
  end

end
