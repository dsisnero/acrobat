require "win32ole"
require_relative "pdoc"
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
    attr_reader :ole_obj

    # the wrapped [PDoc] PDoc object
    attr_reader :pdoc

    # Initialize the
    # @return [App] return an instance of App
    def initialize
      @ole_obj = WIN32OLE.new("AcroExch.App")
      load_constants(@ole_obj)
      @docs = []
    end

    # Runs the adobe app and quits at the end
    # @yield app [App]
    #
    # @example
    #   Acrobat::App.run do |app|
    #     doc = app.open('doc.pdf')
    #     doc.fill_form( city: 'City', state: 'ST')
    #     doc.save_as('filled.pdf')
    #   end
    #
    # @return nil
    def self.run
      the_app = new
      yield the_app
    ensure
      the_app&.quit
      GC.start
      nil
    end

    def self.replace_pages(src, replacement, output_name:, **opts)
      run do |app|
        app.open(src) do |doc|
          doc.replace_pages(replacement, **opts)
          doc.save_as(output_name)
        end
      end
    end

    # Fills the form with updates in a hash
    # @example
    #    Acrobat::App.fill_form(myform.pdf, output_name: 'filled.pdf
    #                                 , update_hash: { name: 'dom', filled_date: 1/20/2013
    # @param doc [String] the String path of a fillable pdf file
    # @param output_name [String] the name of the saved filled pdf file
    # @param update_hash [Hash] the hash with updates
    def self.fill_form(form, output_name:, update_hash:)
      run do |app|
        doc = app.open(form)
        doc.fill_form(update_hash)
        doc.save_as(output_name)
      end
    end

    # show the Adobe Acrobat application
    def show
      ole_obj.Show
    end

    # hide the Adobe Acrobat application
    def hide
      ole_obj.Hide
    end

    # Finds the pdfs in a dir
    # @param dir [String] the directory to find pdfs in
    # @return [Array] of pdf files
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
    # @param dir [String] the path of the directory
    # @param name [String,Nil] the name of the returned pdf file
    #    if the name is nil, the name is "merged.pdf"
    # @param output_dir [String,Nil] the name of the output dir
    #    if the output_dir is nil, the output dir is the dir param
    # return [Boolean] if the merge was successful or not
    def merge_pdfs_in_dir(dir, name: nil, output_dir: nil)
      name || "merged.pdf"
      dir = output_dir || dir
      merge_pdfs(find_pdfs_in_dir(dir))
    end

    # quit the Adobe App.
    #   closes the open adobe documents and quits the program
    # @return nil
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
    # @param file [String #to_path]
    # @return [PDoc] the open file as a Pdoc instance
    def open(file)
      filepath = Pathname(file).expand_path
      raise FileNotFound.new(filepath) unless filepath.file?
      pdoc = WIN32OLE.new("AcroExch.PDDoc")
      is_opened = pdoc.open FileSystemObject.windows_path(filepath)
      doc = PDoc.new(self, pdoc, filepath) if is_opened
      docs << doc if is_opened
      return doc unless block_given?
      begin
        yield doc
      ensure
        doc.close
        nil
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
