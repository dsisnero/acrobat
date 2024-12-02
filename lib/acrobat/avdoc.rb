module Acrobat
  class AvDoc
    attr_reader :ole_obj #: WIN32OLE

    attr_reader :pdoc_name #: String

    def initialize( ole)
      @ole_obj = ole
      @pdoc_name = ole.GetPDDoc.GetFileName
    end

    # Close the AvDoc
    # @rbs save: Boolean -- if true asks to save if false closes without saving
    def close(save = false)
      ole_save = save ? 0 : 1
      @ole_obj.Close(ole_save)
      @ole_obj = nil
      true
    end

  end
end
