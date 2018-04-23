require 'win32ole'
require 'acrobat/app'


module Acrobat
  VERSION = "0.0.8"
end


if $0 ==  __FILE__
  require 'pry'

  app = Acrobat::App.run do |app|

    data = Pathname(__dir__).parent + 'data'
    antenna_form = data + 'faa.6030.17.antenna.pdf'


    doc1 = app.open(antenna_form)
    doc1.show

    fields = {'city' => 'OGD', 'state' => 'Utah',
              'lid' => 'OGD',
              'fac' => 'RTR',
             }
    doc1.fill_form(fields)


    doc1.save_as(name: 'ogd.rtr.pdf', dir: 'tmp')
    binding.pry
    jso = doc1.jso
    jso.show_console
    puts "field count: #{jso.field_count}"
    puts "field names: \n#{jso.field_names}"
    binding.pry



    doc2 = app.open(data + 'faa.6030.17.cm300.uhf.tx.pdf')
    doc2.show
    doc2.fill_form(fields)

    doc1.merge(doc2)
    doc1.save_as(name: 'ogd.merged_antenna_tx.pdf', dir: 'tmp')

  end

end
