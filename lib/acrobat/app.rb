# rbs_inline: enable

require "win32ole"
require_relative "pdoc"
require_relative "avdoc"
require_relative "jso"

module Acrobat
  class FileNotFound < StandardError
    def initialize(path)
      super("File not found: #{path}")
    end
  end
end

module FileSystemObject
  @instance = nil
  def self.instance
    @instance ||= WIN32OLE.new("Scripting.FileSystemObject")
    @instance
  end

  def self.windows_path(path)
    FileSystemObject.instance.GetAbsolutePathname(path.to_s)
  end
end

module ACRO; end

module Acrobat
  class App
    #  [WIN32_OLE] ole_obj
    attr_reader :ole_obj # :WIN32OLE

    def initialize # : Void
      @ole_obj = WIN32OLE.new("AcroExch.App")
      unless @ole_obj
        @ole_obj = WIN32OLE.connect("AcroExch.App")
      end
      load_constants(@ole_obj)
      @docs = []
    end

    def self.close(path)
      name = File.basename(path)
      run do |app|
        av = app.avdocs.find{|a| a.pdoc_name == name}
        av&.close
      end
    end
        


    # Runs the adobe app and quits at the end
    # @example
    #   Acrobat::App.run do |app|
    #     doc = app.open('doc.pdf')
    #     doc.fill_form( city: 'City', state: 'ST')
    #     doc.save_as('filled.pdf')
    #   end
    #
    # @rbs return Void
    def self.run
      the_app = new
      yield the_app
    ensure
      the_app&.quit
      GC.start
    end

    def self.replace_pages(pdf_file, replacement, output_name:, **opts)
      run do |app|
        app.open(pdf_file) do |doc|
          doc.replace_pages(replacement, **opts)
          doc.save_as(output_name)
        end
      end
    end

    # Fills the form with updates in a hash
    # @example
    #    Acrobat::App.fill_form(myform.pdf, output_name: 'filled.pdf
    #                                 , update_hash: { name: 'dom', filled_date: 1/20/2013
    # @rbs pdf_form: String -- the String path of a fillable pdf file
    # @rbs output_name: String -- the name of the saved filled pdf file
    # @rbs update_hash: Hash -- the hash with updates
    def self.fill_form(pdf_form, output_name:, update_hash:)
      run do |app|
        doc = app.open(pdf_form)
        doc.fill_form(update_hash)
        doc.save_as(output_name)
      end
    end

    # @rbs return Array[AvDoc]
    # @rbs &: (AvDoc) -> Void
    def avdocs
      count = ole_obj.GetNumAVDocs
      return [] if count == 0
      return to_enum(:avdocs) unless block_given? 
      av_array = []
      count.times do |i|
        ole = ole_obj.GetAVDoc(i)
        raise "Wrong index for GetNumAVDocs #{i -1}" unless ole
        av = AvDoc.new(ole)
        if block_given?
          yield av
        else
          av_array << av
        end
      end
      av_array unless block_given?
    end
        

    # show the Adobe Acrobat application
    def show # : Void
      ole_obj.Show
    end

    # hide the Adobe Acrobat application
    def hide # : Void
      ole_obj.Hide
    end

    # Finds the pdfs in a dir
    # @rbs dir: String -- the directory to find pdfs in
    # @rbs return Array[Pathname] -- of pdf files
    def find_pdfs_in_dir(dir)
      Pathname.glob(dir + "/*.pdf")
    end

    def merge_pdfs(*pdfs)
      pdf_array = Array(pdfs).flatten
      raise "Not enough pdfs to merge" if pdf_array.size < 2
      first, *rest = pdf_array
      doc = self.open(first)
      rest.each do |path|
        doc.merge(path)
      end
      doc
    end

    # merges the pdfs in directory
    # @rbs dir: String -- the path of the directory
    # @rbs name: String | Nil -- the name of the returned pdf file
    #    if the name is nil, the name is "merged.pdf"
    # @rbs output_dir: String | Nil -- the name of the output dir
    #    if the output_dir is nil, the output dir is the dir param
    # return [Boolean] if the merge was successful or not
    def merge_pdfs_in_dir(dir, name: nil, output_dir: nil)
      name || "merged.pdf"
      dir = output_dir || dir
      merge_pdfs(find_pdfs_in_dir(dir))
    end

    # quit the Adobe App.
    #   closes the open adobe documents and quits the program
    # @rbs return nil
    def quit
      begin
        docs.each { |d| d.close }
        ole_obj.Exit
      rescue
        return nil
      end
      nil
    end

    attr_reader :docs

    # open the file.
    # @rbs file: String | Pathname -- #to_path
    # @rbs &: (PDoc) -> Nil
    # @rbs return PDoc -- the open file as a Pdoc instance
    def open(file)
      doc = PDoc.from_path(file)
      docs << doc
      return doc unless block_given?
      begin
        yield doc
      ensure
        doc.close
      end
    end

    def form
      Form.new(self, WIN32OLE.new("AFormAut.App"))
    end

    private

    def load_constants(ole_obj)
      WIN32OLE.const_load(ole_obj, ACRO) unless ACRO.constants.size > 0
    end
  end
end
