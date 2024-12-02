require_relative "../test_helper"
require "acrobat"
class TestAcrobatApp < Minitest::Test
  def setup
    @app = Acrobat::App.new
  end

  def test_initialize
    refute_nil(@app.ole_obj)
    assert_instance_of(WIN32OLE, @app.ole_obj)
    assert_empty(@app.docs)
  end

  def test_open
    pdf_path = File.expand_path(TEST_FILES + "6030.17.antenna.pdf")
    doc = @app.open(pdf_path)
    assert_instance_of(Acrobat::PDoc, doc)
    assert_equal(1, @app.docs.size)
    assert_equal(doc, @app.docs.first)
  end

  def test_open_with_invalid_file
    assert_raises(Acrobat::FileNotFound) { @app.open("invalid/path/to/file.pdf") }
  end

  def test_merge_pdfs
    pdf1_path = File.expand_path(TEST_FILES / "6030.17.antenna.pdf")
    pdf2_path = File.expand_path(TEST_FILES / "6030.17.cm300.tx.pdf")
    page_count1 =  @app.open(pdf1_path){|f| f.page_count}
    page_count2 = @app.open(pdf2_path){|f| f.page_count}
        merged_doc = @app.merge_pdfs(pdf1_path, pdf2_path)
    merged_doc.save_as("merged_ant_tx.pdf")
    assert_instance_of(Acrobat::PDoc, merged_doc)
    assert_equal(page_count1 + page_count2 , merged_doc.page_count)
  end

  def test_find_pdfs_in_dir
    pdf_dir = File.expand_path(TEST_FILES)
    pdfs = @app.find_pdfs_in_dir(pdf_dir)
    assert_equal(2, pdfs.size)
  end

  def test_merge_pdfs_in_dir
    pdf_dir = File.expand_path(TEST_FILES)
    merged_doc = @app.merge_pdfs_in_dir(pdf_dir)

    assert_instance_of(Acrobat::PDoc, merged_doc)
  end

  def test_quit
    skip
    pdf_path = File.expand_path(TEST_FILES + "6030.17.antenna.pdf")
    @app.open(pdf_path)
    @app.quit

    assert_empty(@app.docs)
  end
end
