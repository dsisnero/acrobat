# frozen_string_literal: true

require "test_helper"

class TestAcrobat < Minitest::Test
  def test_that_it_has_a_version_number
    refute_nil ::Acrobat::VERSION
  end
end
