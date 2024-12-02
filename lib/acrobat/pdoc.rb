module Acrobat
  class PDoc
    attr_reader :ole_obj
    
    attr_reader  :path #: Path | Nil

    def self.from_path(path)
      filepath = Pathname(path).expand_path
      raise FileNotFound.new(filepath) unless filepath.file?
      pdoc = WIN32OLE.new("AcroExch.PDDoc")
      result = pdoc.open FileSystemObject.windows_path(filepath)
      raise "Error opening Acrobat::Pdoc from #{filepath}" if result == 0
      new(pdoc, path)
    end

    def initialize(ole, path = nil)
      @ole_obj = ole
      @path = path
    end

    def show(name = nil)
      name ||= ole_obj.GetFileName
      ole_obj.OpenAVDoc(name)
    end

    # @rbs return Fixnum -- the number of pages in the pdf
    def page_count
      ole_obj.GetNumPages()
    end

    def last_page # : Fixnum
      page_count - 1
    end

    # merges the doc to the
    #   @rbs doc: PDoc|String|Pathname --  an open PDoc to merge or filename
    # @rbs return Boolean -- whether the doc was merged correctly
    def merge(doc)
      src = open_ole_pdoc(doc)
      ole_insert_pages(src, self_start_page: last_page, merge_doc_page_start: 0,
        number_pages: src.page_count, bookmarks: true)
    end

    # opens and/or returns PDoc
    #   @rbs doc: String|Pathname|Pdoc  -- the String path of a pdf file
    def open_ole_pdoc(doc) # : Pdoc
      case doc
      when PDoc
        doc
      when String, Pathname
        self.class.from_path(doc)
      end
    end

    def fill_and_save(results, name: nil, dir: nil)
      fill_form(results)
      is_saved = save_as(name: name, dir: dir)
      puts "saved file: %s" % [dir + name] if is_saved
      true
    end

    def prepend_pages(src_path:, src_page_start: 1, src_page_end: nil)
      insert_pages(insert_after: 0, src_path: src_path, src_page_start: src_page_start, src_page_end: src_page_end)
    end

    # returns [Pathname] of d
    #   @rbs dir: [String, Nil] the String path
    # @rbs return dir [Pathname] Pathname of dir or of working directory
    def default_dir(dir)
      Pathname(dir || Pathname.getw)
    end

    def save_as(name, dir: nil)
      name ||= path.basename
      dir = Pathname(dir || Pathname.getwd)
      dir.mkpath
      windows_path = FileSystemObject.windows_path(dir + name)
      ole_obj.save(ACRO::PDSaveFull | ACRO::PDSaveCopy, windows_path)
    end

    def name
      ole_obj.GetFileName
    end

    def close # : Bool -- true if closes successfully
      result = ole_obj.Close
      result == -1
    rescue
      false
    end

    def replace_pages(doc, start: 0, replace_start: 0, num_of_pages: 1, merge_annotations: true)
      src = open_ole_pdoc(doc)

      ole_obj.ReplacePages(start, src.ole_obj, replace_start, num_of_pages, merge_annotations)
    end

    # return the instance of JSO object
    # @rbs return [Jso] a WIN32OLE wrapped Jso object 'javascript object'
    def jso
      @jso ||= Jso.new(self, ole_obj.GetJSObject)
    end

    # return the field_names of the pdf form
    def field_names
      jso.field_names
    end

    def fill_form(results)
      jso.fill_form(results)
    end

    protected

    def ole_insert_pages(merge_doc, self_start_page: 0, merge_doc_page_start: 0,
      number_pages: nil, bookmarks: true)
      bookmarks_ole = bookmarks ? 1 : 0
      result = ole_obj.InsertPages(self_start_page, merge_doc.ole_obj, merge_doc_page_start, number_pages, bookmarks_ole)
      if result == 0
        raise "ole.InsertPages unable to insert merge doc #{merge_doc.name} into #{name}"
      end
      true
    end

    def ole_replace_pages(merge_doc, self_start_page: nil, merge_doc_page_start: nil,
      number_pages: nil, bookmarks: true)
      self_start_page_ole = self_start_page ? self_start_page - 1 : page_count - 1
      page_start_ole = merge_doc_page_start ? merge_doc_page_start - 1 : merge_doc.GetNumPages - 1
      number_pages_ole = number_pages || merge_doc
      bookmarks_ole = bookmarks ? 1 : 0
      result = ole_obj.ReplacePages(self_start_page_ole, merge_doc.ole, page_start_ole, number_pages_ole, bookmarks_ole)
    end

    def insert_pages(src:, insert_after: nil, src_page_start: nil, src_page_end: nil)
      insert_hash = {"nPage" => insert_after - 1}
      insert_hash["nStart"] = src_page_start + 1 if src_page_start
      insert_hash["nEnd"] = src_page_end + 1 if src_page_end
      ole_obj.InsertPages(**insert_hash)
    end

    def merge_pdoc(doc, **options)
      if options
        start = options[:start]
        start_page
        options[:pages]
        ole_obj.InsertPages(start, doc.ole_obj, 0, doc.page_count, true)
      else
        ole_obj.InsertPages(page_count - 1, doc.ole_obj, 0, doc.page_count, true)

      end
    rescue
      false
    end
  end
end
