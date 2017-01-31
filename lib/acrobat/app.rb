require 'pry'
require  'win32ole'

module FileSystemObject

  @instance = nil
  def FileSystemObject.instance
    unless @instance
      @instance = WIN32OLE.new('Scripting.FileSystemObject')
    end
    return @instance
  end

  def FileSystemObject.windows_path(path)
    FileSystemObject.instance.GetAbsolutePathname(path.to_s)
  end

end

module ACRO; end

module Acrobat

  class App


    attr_reader :ole_obj, :pdoc

    def initialize()
      @ole_obj = WIN32OLE.new('AcroExch.App')
      load_constants(@ole_obj)
      @ole_obj
      @docs = []
      self
    end

    # Runs the adobe app and quits at the end
    # @yield app [App]
    #
    # Acrobat::App.run do |app|
    #    doc = app.open('doc.pdf')
    #    doc.fill_form( city: 'City', state: 'ST')
    #    doc.save_as('filled.pdf')
    # end
    #
    # @return nil
    def self.run
      begin
        the_app = new
        yield the_app
      ensure
        the_app.quit unless the_app.nil?
        the_app = nil
        GC.start
        nil
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
      Pathname.glob( dir + '*.pdf')
    end

    # merges the pdfs in directory
    # @param dir [String] the path of the directory
    # @param name [String,Nil] the name of the returned pdf file
    #    if the name is nil, the name is "merged.pdf"
    # @param output_dir [String,Nil] the name of the output dir
    #    if the output_dir is nil, the output dir is the dir param
    # return [Boolean] if the merge was successful or not
    def merge_pdfs(dir, name: nil , output_dir: nil)
      name = lname || "merged.pdf"
      dir = output_dir || dir
      pdfs = Pathname.glob( dir + '*.pdf')
      raise 'Not enough pdfs to merge' if pdfs.size < 2
      first, *rest = pdfs
      open(first) do |pdf|
        rest.each do |path|
          open(path) do |pdfnext|
            pdf.merge_doc(pdfnext)
          end
        end

        pdf.save_as(name: name, dir: dir)
      end
    end

    # quit the Adobe App.
    #   closes the open adobe documents and quits the program
    # @return nil
    def quit
      begin
        docs.each{ |d| d.close}
        ole_obj.Exit
      rescue
        return nil
      end
      nil
    end

    def docs
      @docs
    end

    # open the file.
    # @param file [String #to_path]
    # @return [PDoc] the open file as a Pdoc instance
    def open(file)
      filepath = Pathname(file).expand_path
      raise FileNotFound unless filepath.file?
      pdoc = WIN32OLE.new('AcroExch.PDDoc')
      is_opened = pdoc.open FileSystemObject.windows_path(filepath)
      doc = PDoc.new(self, pdoc, filepath) if is_opened
      docs << doc if is_opened
      return doc unless block_given?
      begin
        yield doc
      ensure
        doc.close unless doc.closed?
        doc = nil
      end
    end


    def form
      Form.new(self,WIN32OLE.new("AFormAut.App"))
    end

    private

    def load_constants(ole_obj)
      WIN32OLE.const_load(ole_obj, ACRO) unless ACRO.constants.size > 0
    end


  end

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

    def update_form(results)
      form.update(results)
    end

    def update_and_save(results,name: nil, dir: nil)
      update_form(results)
      is_saved = save_as(name: name, dir: dir)
      puts "saved file: %s" % [dir + name] if is_saved
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
      ole_obj.Close
    end


    def jso
      @jso ||= Jso.new(ole_obj.GetJSObject)
    end

    def field_names
      jso.field_names
    end

    def fill_form(results)
      jso.fill_form(results)
    end

    protected
    def merge_pdoc(doc)
      ole_obj.InsertPages(page_count - 1, doc.ole_obj, 0, doc.page_count, true)
    end

  end

  class Jso

    attr_reader :ole_obj

    def initialize(ole)
      @ole_obj = ole
    end

    def find_field(name_or_number)
      case name_or_number
      when String,Symbol
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

    def field_names
      result = []
      0.upto(field_count -1) do |i|
        result << ole_obj.getNthFieldName(i)
      end
      result
    end

    def set_field(name,value)
      field = find_field(name)
      field.Value = value if field
    end

    def get_field(name)
      field = find_field(name)
      field.Value if field
    end

    def field_count
      ole_obj.numFields
    end

    def clear_form
      ole_obj.resetForm
    end

    def fill_form(hash)
      clear_form
      hash.each do |k,v|
        set_field(k,v)
      end
    end

  end


  class Form

    attr_reader :app, :ole_obj

    def initialize(app, ole_obj)
      @app = app
      @ole_obj = ole_obj
    end

    def ole_fields
      @ole_obj.fields
    end

    def [](name)
      ole = get_name(name)
      ole.value if ole
    end

    def get_name(name)
      ole_obj.fields.Item(name.to_s) rescue nil
    end

    def []=(name,value)
      ole = get_name(name.to_s)
      ole.Value = value if ole
    end

    def update(hash)
      hash.each do |k,v|
        self[k.to_s] = v
      end

    end

    def to_hash
      result = {}
      ole_obj.fields.each do |fld|
        result[fld.Name] = fld.value
      end
      result
    end
  end



end

if $0 ==  __FILE__
  require 'pry'

  app = Acrobat::App.run do |app|

    data = Pathname(__dir__).parent.parent + 'data'
    antenna_form = data + 'faa.6030.17.antenna.pdf'


    doc1 = app.open(antenna_form)
    doc1.show

    fields = {'city' => 'OGD', 'state' => 'Utah',
              'lid' => 'OGD',
              'fac' => 'RTR',
             }
    doc1.fill_form(fields)


    doc1.save_as(name: 'ogd.rtr.pdf', dir: 'tmp')




    doc2 = app.open(data + 'faa.6030.17.cm300.uhf.tx.pdf')
    doc2.show
    doc2.fill_form(fields)

    doc1.merge(doc2)
    doc1.save_as(name: 'ogd.merged_antenna_tx.pdf', dir: 'tmp')

  end

end



# ublic Sub Main()
#     Dim AcroApp As Acrobat.CAcroApp
#     Dim AvDoc As Acrobat.CAcroAVDoc
#     Dim fcount As Long
#     Dim sFieldName As String
#     Dim Field As AFORMAUTLib.Field
#     Dim Fields As AFORMAUTLib.Fields
#     Dim AcroForm As AFORMAUTLib.AFormApp

#     Set AcroApp = CreateObject("AcroExch.App")
#     Set AvDoc = CreateObject("AcroExch.AVDoc")

#     If AvDoc.Open("C:\test\testform.pdf", "") Then
#         AcroApp.Show
#         Set AcroForm = CreateObject("AFormAut.App")
#         Set Fields = AcroForm.Fields
#         fcount = Fields.Count
#         MsgBox fcount
#         For Each Field In Fields
#             sFieldName = Field.Name
#             MsgBox sFieldName
#         Next Field
#     Else
#         MsgBox "failure"
#     End If
#     AcroApp.Exit
#     Set AcroApp = Nothing
#     Set AvDoc = Nothing
#     Set Field = Nothing
#     Set Fields = Nothing
#     Set AcroForm = Nothing
