require "acrobat/app"
require "pry"
Acrobat::App.run do |app|
  dir = Pathname(__dir__)
  form_path = dir + "6030.17.antenna.pdf"
  doc = app.open(form_path)
  binding.pry
  doc.merge(dir + "6030.17.cm300.tx.pdf")
  doc.show
  doc.save_as(name: "example.merged.pdf", dir: dir.parent + "tmp")
end
