require_relative 'test_helper'


require "acrobat"



module Acrobat

  describe App do


    describe 'new' do

      it 'works' do
        assert true
      end


      it 'can be created' do
        app = App.new
        app.must_be_instance_of App
      end

      it 'has an ole_obj' do
        app = App.new
        app.ole_obj.must_be_instance_of WIN32OLE
       end

    end


  end


end
