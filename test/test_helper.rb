# frozen_string_literal: true

$LOAD_PATH.unshift File.expand_path("../lib", __dir__)
require "acrobat"

require "minitest/autorun"

TEST_FILES = Pathname(__dir__) + "data"
