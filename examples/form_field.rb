require "acrobat"
require "pry"

Acrobat::App.new
# app.show


Acrobat::App.run do |app|
  dir = Pathname(__dir__)
  form_path = dir + "6030.17.antenna.pdf"
  puts "Working on #{form_path}"
  doc1 = app.open(form_path)
  doc1.show
  fields = {"city" => "Salt Lake City",
            "state" => "Utah",
            "lid" => "SLCB",
            "fac" => "RTR",
            :cost_center => "1234"}
  doc1.fill_form(fields)
  doc1.save_as(name: "slcb.example.pdf", dir: dir.parent + "tmp")
end
